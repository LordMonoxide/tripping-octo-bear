VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEditor_Events 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Events Editor"
   ClientHeight    =   9600
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   640
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1018
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraLabeling 
      Caption         =   "Labeling Variables and Switches"
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   495
      Begin VB.Frame fraRenaming 
         Caption         =   "Renaming Variable/Switch"
         Height          =   7455
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   8895
         Begin VB.Frame fraRandom 
            Caption         =   "Editing Variable/Switch"
            Height          =   2295
            Index           =   10
            Left            =   1920
            TabIndex        =   18
            Top             =   2640
            Width           =   5055
            Begin VB.CommandButton cmdRename_Ok 
               Caption         =   "Ok"
               Height          =   375
               Left            =   2280
               TabIndex        =   21
               Top             =   1800
               Width           =   1215
            End
            Begin VB.CommandButton cmdRename_Cancel 
               Caption         =   "Cancel"
               Height          =   375
               Left            =   3720
               TabIndex        =   20
               Top             =   1800
               Width           =   1215
            End
            Begin VB.TextBox txtRename 
               Height          =   375
               Left            =   120
               TabIndex        =   19
               Top             =   720
               Width           =   4815
            End
            Begin VB.Label lblEditing 
               Caption         =   "Naming Variable #1"
               Height          =   375
               Left            =   120
               TabIndex        =   22
               Top             =   360
               Width           =   4815
            End
         End
      End
      Begin VB.CommandButton cmbLabel_Ok 
         Caption         =   "OK"
         Height          =   375
         Left            =   5760
         TabIndex        =   166
         Top             =   7200
         Width           =   1455
      End
      Begin VB.CommandButton cmdLabel_Cancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   7320
         TabIndex        =   165
         Top             =   7200
         Width           =   1455
      End
      Begin VB.ListBox lstVariables 
         Height          =   5520
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   4335
      End
      Begin VB.ListBox lstSwitches 
         Height          =   5520
         Left            =   4560
         TabIndex        =   25
         Top             =   720
         Width           =   4455
      End
      Begin VB.CommandButton cmdRenameVariable 
         Caption         =   "Rename Variable"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   6240
         Width           =   4335
      End
      Begin VB.CommandButton cmdRenameSwitch 
         Caption         =   "Rename Switch"
         Height          =   375
         Left            =   4560
         TabIndex        =   23
         Top             =   6240
         Width           =   4455
      End
      Begin VB.Label lblRandomLabel 
         Alignment       =   2  'Center
         Caption         =   "Player Variables"
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label lblRandomLabel 
         Alignment       =   2  'Center
         Caption         =   "Player Switches"
         Height          =   255
         Index           =   36
         Left            =   4560
         TabIndex        =   27
         Top             =   480
         Width           =   4455
      End
   End
   Begin VB.Frame fraCommands 
      Caption         =   "Add command"
      Height          =   4695
      Left            =   9480
      TabIndex        =   112
      Top             =   4800
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton cmdAddOk 
         Caption         =   "Close"
         Height          =   375
         Left            =   120
         TabIndex        =   113
         Top             =   4200
         Width           =   5535
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3735
         Left            =   120
         TabIndex        =   114
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   6588
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "1"
         TabPicture(0)   =   "frmEditor_Event.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmdCommand(12)"
         Tab(0).Control(1)=   "cmdCommand(11)"
         Tab(0).Control(2)=   "cmdCommand(10)"
         Tab(0).Control(3)=   "cmdCommand(0)"
         Tab(0).Control(4)=   "cmdCommand(1)"
         Tab(0).Control(5)=   "cmdCommand(2)"
         Tab(0).Control(6)=   "cmdCommand(3)"
         Tab(0).Control(7)=   "cmdCommand(4)"
         Tab(0).Control(8)=   "cmdCommand(5)"
         Tab(0).Control(9)=   "cmdCommand(6)"
         Tab(0).Control(10)=   "cmdCommand(7)"
         Tab(0).Control(11)=   "cmdCommand(8)"
         Tab(0).Control(12)=   "cmdCommand(9)"
         Tab(0).Control(13)=   "cmdCommand(13)"
         Tab(0).Control(14)=   "cmdCommand(14)"
         Tab(0).Control(15)=   "cmdCommand(15)"
         Tab(0).Control(16)=   "cmdCommand(16)"
         Tab(0).Control(17)=   "cmdCommand(17)"
         Tab(0).ControlCount=   18
         TabCaption(1)   =   "2"
         TabPicture(1)   =   "frmEditor_Event.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "cmdCommand(18)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "cmdCommand(19)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "cmdCommand(20)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cmdCommand(21)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "cmdCommand(22)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).ControlCount=   5
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Event Graphic"
            Height          =   375
            Index           =   22
            Left            =   240
            TabIndex        =   206
            Top             =   2400
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change Exp"
            Height          =   375
            Index           =   17
            Left            =   -71400
            TabIndex        =   205
            Top             =   2880
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Spawn NPC"
            Height          =   375
            Index           =   21
            Left            =   240
            TabIndex        =   204
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change PK"
            Height          =   375
            Index           =   16
            Left            =   -71400
            TabIndex        =   198
            Top             =   2400
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Open/Close ev"
            Height          =   375
            Index           =   20
            Left            =   240
            TabIndex        =   196
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Custom script"
            Height          =   375
            Index           =   19
            Left            =   240
            TabIndex        =   188
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Set Access"
            Height          =   375
            Index           =   18
            Left            =   240
            TabIndex        =   184
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change skill"
            Height          =   375
            Index           =   15
            Left            =   -71400
            TabIndex        =   164
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Conditional branch"
            Height          =   375
            Index           =   14
            Left            =   -71400
            TabIndex        =   158
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Chatbubble"
            Height          =   375
            Index           =   13
            Left            =   -71400
            TabIndex        =   135
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "GoTo"
            Height          =   375
            Index           =   9
            Left            =   -73080
            TabIndex        =   127
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Warp"
            Height          =   375
            Index           =   8
            Left            =   -73080
            TabIndex        =   126
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Play animation"
            Height          =   375
            Index           =   7
            Left            =   -73080
            TabIndex        =   125
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change Level"
            Height          =   375
            Index           =   6
            Left            =   -73080
            TabIndex        =   124
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change items"
            Height          =   375
            Index           =   5
            Left            =   -74760
            TabIndex        =   123
            Top             =   2880
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Open Bank"
            Height          =   375
            Index           =   4
            Left            =   -74760
            TabIndex        =   122
            Top             =   2400
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Open Shop"
            Height          =   375
            Index           =   3
            Left            =   -74760
            TabIndex        =   121
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Exit event"
            Height          =   375
            Index           =   2
            Left            =   -74760
            TabIndex        =   120
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Show choices"
            Height          =   375
            Index           =   1
            Left            =   -74760
            TabIndex        =   119
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Show message"
            Height          =   375
            Index           =   0
            Left            =   -74760
            TabIndex        =   118
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change Switch"
            Height          =   375
            Index           =   10
            Left            =   -73080
            TabIndex        =   117
            Top             =   2400
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change Variable"
            Height          =   375
            Index           =   11
            Left            =   -73080
            TabIndex        =   116
            Top             =   2880
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Add Text"
            Height          =   375
            Index           =   12
            Left            =   -71400
            TabIndex        =   115
            Top             =   480
            Width           =   1455
         End
      End
   End
   Begin VB.Frame fraEditCommand 
      Caption         =   "Edit Command"
      Height          =   4695
      Left            =   9480
      TabIndex        =   44
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Frame fraChangeGraphic 
         Caption         =   "Change event graphic"
         Height          =   3735
         Left            =   120
         TabIndex        =   207
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.HScrollBar scrlChangeGraphic 
            Height          =   255
            Left            =   1560
            Max             =   2
            TabIndex        =   213
            Top             =   1200
            Width           =   1455
         End
         Begin VB.HScrollBar scrlChangeGraphicX 
            Height          =   255
            Left            =   1560
            Max             =   255
            TabIndex        =   210
            Top             =   360
            Width           =   3855
         End
         Begin VB.HScrollBar scrlChangeGraphicY 
            Height          =   255
            Left            =   1560
            Max             =   255
            TabIndex        =   209
            Top             =   720
            Width           =   3855
         End
         Begin VB.ComboBox cmbChangeGraphicType 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":0038
            Left            =   3120
            List            =   "frmEditor_Event.frx":0045
            TabIndex        =   208
            Text            =   "cmbChangeGraphicType"
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label lblChangeGraphic 
            Caption         =   "Graphic#: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   214
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lblChangeGraphicX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   212
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblChangeGraphicY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   211
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdEditOk 
         Caption         =   "Close"
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   4200
         Width           =   5535
      End
      Begin VB.Frame fraOpenEvent 
         Caption         =   "Open/Close event"
         Height          =   3735
         Left            =   120
         TabIndex        =   189
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.ComboBox cmbOpenEventType 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":007C
            Left            =   3120
            List            =   "frmEditor_Event.frx":0089
            TabIndex        =   197
            Text            =   "cmbOpenEventType"
            Top             =   1200
            Width           =   2295
         End
         Begin VB.OptionButton optOpenEventType 
            Caption         =   "Close"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   195
            Top             =   1200
            Width           =   855
         End
         Begin VB.OptionButton optOpenEventType 
            Caption         =   "Open"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   194
            Top             =   1200
            Width           =   735
         End
         Begin VB.HScrollBar scrlOpenEventY 
            Height          =   255
            Left            =   1560
            Max             =   255
            TabIndex        =   192
            Top             =   720
            Width           =   3855
         End
         Begin VB.HScrollBar scrlOpenEventX 
            Height          =   255
            Left            =   1560
            Max             =   255
            TabIndex        =   190
            Top             =   360
            Width           =   3855
         End
         Begin VB.Label lblOpenEventY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   193
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblOpenEventX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   191
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame fraCustomScript 
         Caption         =   "Execute Custom Script"
         Height          =   3735
         Left            =   120
         TabIndex        =   185
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.HScrollBar scrlCustomScript 
            Height          =   255
            Left            =   1560
            Max             =   255
            TabIndex        =   186
            Top             =   360
            Width           =   3855
         End
         Begin VB.Label lblCustomScript 
            Caption         =   "Case: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   187
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame fraSetAccess 
         Caption         =   "Set Access"
         Height          =   3735
         Left            =   120
         TabIndex        =   182
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.ComboBox cmbSetAccess 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":00C0
            Left            =   240
            List            =   "frmEditor_Event.frx":00D3
            Style           =   2  'Dropdown List
            TabIndex        =   183
            Top             =   360
            Width           =   5055
         End
      End
      Begin VB.Frame fraChangeExp 
         Caption         =   "Change Experience"
         Height          =   3735
         Left            =   120
         TabIndex        =   176
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.OptionButton optExpAction 
            Caption         =   "Subtract"
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   181
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton optExpAction 
            Caption         =   "Add"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   180
            Top             =   960
            Width           =   735
         End
         Begin VB.OptionButton optExpAction 
            Caption         =   "Set"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   179
            Top             =   960
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.HScrollBar scrlChangeExp 
            Height          =   255
            Left            =   120
            Max             =   32000
            TabIndex        =   177
            Top             =   600
            Width           =   3735
         End
         Begin VB.Label lblChangeExp 
            Caption         =   "Exp: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   178
            Top             =   360
            Width           =   3735
         End
      End
      Begin VB.Frame fraChangeLevel 
         Caption         =   "Change Level"
         Height          =   3735
         Left            =   120
         TabIndex        =   167
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.OptionButton optLevelAction 
            Caption         =   "Set"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   172
            Top             =   960
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optLevelAction 
            Caption         =   "Add"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   171
            Top             =   960
            Width           =   735
         End
         Begin VB.OptionButton optLevelAction 
            Caption         =   "Subtract"
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   170
            Top             =   960
            Width           =   975
         End
         Begin VB.HScrollBar scrlChangeLevel 
            Height          =   255
            Left            =   120
            TabIndex        =   168
            Top             =   600
            Width           =   3855
         End
         Begin VB.Label lblChangeLevel 
            Caption         =   "Level: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   169
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame fraChangeVariable 
         Caption         =   "Change Variable"
         Height          =   3735
         Left            =   120
         TabIndex        =   85
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.TextBox txtVariableData 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   3960
            TabIndex        =   93
            Text            =   "0"
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtVariableData 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   2160
            TabIndex        =   92
            Text            =   "0"
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtVariableData 
            Height          =   285
            Index           =   0
            Left            =   1920
            TabIndex        =   91
            Text            =   "0"
            Top             =   960
            Width           =   3495
         End
         Begin VB.OptionButton optVariableAction 
            Caption         =   "Random"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   90
            Top             =   1560
            Width           =   1095
         End
         Begin VB.OptionButton optVariableAction 
            Caption         =   "Subtract"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   89
            Top             =   1200
            Width           =   1095
         End
         Begin VB.OptionButton optVariableAction 
            Caption         =   "Add"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   88
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton optVariableAction 
            Caption         =   "Set"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   87
            Top             =   720
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.ComboBox cmbVariable 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   86
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "High:"
            Height          =   255
            Index           =   37
            Left            =   3480
            TabIndex        =   96
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Low:"
            Height          =   255
            Index           =   13
            Left            =   1680
            TabIndex        =   95
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Variable:"
            Height          =   255
            Index           =   12
            Left            =   360
            TabIndex        =   94
            Top             =   240
            Width           =   3855
         End
      End
      Begin VB.Frame fraMenu 
         Caption         =   "Show choices"
         Height          =   3735
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.HScrollBar scrlMenuOptDest 
            Height          =   255
            Left            =   240
            Max             =   10
            Min             =   1
            TabIndex        =   81
            Top             =   3360
            Value           =   1
            Width           =   5175
         End
         Begin VB.TextBox txtMenuOptText 
            Height          =   285
            Left            =   1440
            TabIndex        =   80
            Top             =   2760
            Width           =   3855
         End
         Begin VB.CommandButton cmdRemoveMenuOption 
            Caption         =   "Remove"
            Height          =   375
            Left            =   3960
            TabIndex        =   79
            Top             =   2280
            Width           =   1455
         End
         Begin VB.CommandButton cmdModifyMenuOption 
            Caption         =   "Modify"
            Height          =   375
            Left            =   2040
            TabIndex        =   78
            Top             =   2280
            Width           =   1335
         End
         Begin VB.CommandButton cmdAddMenuOption 
            Caption         =   "Add"
            Height          =   375
            Left            =   120
            TabIndex        =   77
            Top             =   2280
            Width           =   1335
         End
         Begin VB.ListBox lstMenuOptions 
            Height          =   1035
            Left            =   120
            TabIndex        =   76
            Top             =   1200
            Width           =   5295
         End
         Begin VB.TextBox txtMenuQuery 
            Height          =   645
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   75
            Top             =   480
            Width           =   5325
         End
         Begin VB.Label lblMenuOptDest 
            Caption         =   "Destination: 1"
            Height          =   255
            Left            =   240
            TabIndex        =   84
            Top             =   3120
            Width           =   5175
         End
         Begin VB.Label Label5 
            Caption         =   "Option Text:"
            Height          =   255
            Left            =   240
            TabIndex        =   83
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Menu Query:"
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame fraGiveItem 
         Caption         =   "Change items"
         Height          =   3735
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.HScrollBar scrlGiveItemAmount 
            Height          =   255
            Left            =   120
            Max             =   250
            Min             =   1
            TabIndex        =   71
            Top             =   1080
            Value           =   1
            Width           =   5295
         End
         Begin VB.HScrollBar scrlGiveItemID 
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   480
            Width           =   5295
         End
         Begin VB.OptionButton optItemOperation 
            Caption         =   "Take item"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   69
            Top             =   1440
            Width           =   1455
         End
         Begin VB.OptionButton optItemOperation 
            Caption         =   "Change item"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   68
            Top             =   1440
            Width           =   1455
         End
         Begin VB.OptionButton optItemOperation 
            Caption         =   "Give Item"
            Height          =   255
            Index           =   2
            Left            =   3960
            TabIndex        =   67
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label lblGiveItemAmount 
            Caption         =   "Amount: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   840
            Width           =   5295
         End
         Begin VB.Label lblGiveItemID 
            Caption         =   "Item: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.Frame fraOpenShop 
         Caption         =   "Open Shop"
         Height          =   3735
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.HScrollBar scrlOpenShop 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   64
            Top             =   480
            Value           =   1
            Width           =   5295
         End
         Begin VB.Label lblOpenShop 
            Caption         =   "Open Shop: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.Frame fraMapWarp 
         Caption         =   "Map Warp"
         Height          =   3735
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.HScrollBar scrlWarpY 
            Height          =   255
            Left            =   120
            Max             =   250
            TabIndex        =   59
            Top             =   1680
            Value           =   1
            Width           =   5295
         End
         Begin VB.HScrollBar scrlWarpX 
            Height          =   255
            Left            =   120
            Max             =   250
            TabIndex        =   58
            Top             =   1080
            Value           =   1
            Width           =   5295
         End
         Begin VB.HScrollBar scrlWarpMap 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   57
            Top             =   480
            Value           =   1
            Width           =   5295
         End
         Begin VB.Label lblWarpY 
            Caption         =   "Y: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   1440
            Width           =   5295
         End
         Begin VB.Label lblWarpX 
            Caption         =   "X: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   840
            Width           =   5295
         End
         Begin VB.Label lblWarpMap 
            Caption         =   "Map: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.Frame fraChatbubble 
         Caption         =   "Show Chatbubble"
         Height          =   3735
         Left            =   120
         TabIndex        =   128
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.OptionButton optChatBubbleTarget 
            Caption         =   "NPC"
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   131
            Top             =   1440
            Width           =   735
         End
         Begin VB.ComboBox cmbChatBubbleTarget 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":0116
            Left            =   120
            List            =   "frmEditor_Event.frx":0118
            Style           =   2  'Dropdown List
            TabIndex        =   130
            Top             =   1800
            Visible         =   0   'False
            Width           =   5295
         End
         Begin VB.TextBox txtChatbubbleText 
            Height          =   1005
            Left            =   1680
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   129
            Top             =   360
            Width           =   3735
         End
         Begin VB.OptionButton optChatBubbleTarget 
            Caption         =   "Player"
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   132
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Chatbubble Text:"
            Height          =   255
            Index           =   38
            Left            =   120
            TabIndex        =   134
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Target Type:"
            Height          =   255
            Index           =   39
            Left            =   120
            TabIndex        =   133
            Top             =   1440
            Width           =   1335
         End
      End
      Begin VB.Frame fraPlayerText 
         Caption         =   "Show Message"
         Height          =   3735
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.TextBox txtPlayerText 
            Height          =   3015
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   55
            Top             =   240
            Width           =   5355
         End
      End
      Begin VB.Frame fraAddText 
         Caption         =   "Add Text"
         Height          =   3735
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.TextBox txtAddText_Text 
            Height          =   1815
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   50
            Top             =   240
            Width           =   5295
         End
         Begin VB.HScrollBar scrlAddText_Colour 
            Height          =   255
            Left            =   120
            Max             =   18
            TabIndex        =   49
            Top             =   2400
            Width           =   5295
         End
         Begin VB.OptionButton optChannel 
            Caption         =   "Player"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   48
            Top             =   2760
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optChannel 
            Caption         =   "Map"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   47
            Top             =   2760
            Width           =   1095
         End
         Begin VB.OptionButton optChannel 
            Caption         =   "Global"
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   46
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label lblAddText_Colour 
            Caption         =   "Colour: Black"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   2160
            Width           =   3255
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Channel:"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   51
            Top             =   2760
            Width           =   1575
         End
      End
      Begin VB.Frame fraGoTo 
         Caption         =   "GoTo"
         Height          =   3735
         Left            =   120
         TabIndex        =   109
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.HScrollBar scrlGOTO 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   110
            Top             =   480
            Value           =   1
            Width           =   5295
         End
         Begin VB.Label lblGOTO 
            Caption         =   "Goto: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.Frame fraAnimation 
         Caption         =   "Animation"
         Height          =   3735
         Left            =   120
         TabIndex        =   102
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.HScrollBar scrlPlayAnimationY 
            Height          =   255
            Left            =   120
            Max             =   250
            Min             =   -1
            TabIndex        =   105
            Top             =   1680
            Value           =   1
            Width           =   5295
         End
         Begin VB.HScrollBar scrlPlayAnimationX 
            Height          =   255
            Left            =   120
            Max             =   250
            Min             =   -1
            TabIndex        =   104
            Top             =   1080
            Value           =   1
            Width           =   5295
         End
         Begin VB.HScrollBar scrlPlayAnimationAnim 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   103
            Top             =   480
            Value           =   1
            Width           =   5295
         End
         Begin VB.Label lblPlayAnimationY 
            Caption         =   "Y: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Top             =   1440
            Width           =   5295
         End
         Begin VB.Label lblPlayAnimationX 
            Caption         =   "X: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   107
            Top             =   840
            Width           =   5295
         End
         Begin VB.Label lblPlayAnimationAnim 
            Caption         =   "Animation: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   106
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.Frame fraChangeSwitch 
         Caption         =   "Change switch"
         Height          =   3735
         Left            =   120
         TabIndex        =   97
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.ComboBox cmbPlayerSwitchSet 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":011A
            Left            =   1200
            List            =   "frmEditor_Event.frx":0124
            Style           =   2  'Dropdown List
            TabIndex        =   99
            Top             =   795
            Width           =   4215
         End
         Begin VB.ComboBox cmbSwitch 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   98
            Top             =   360
            Width           =   4215
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Switch:"
            Height          =   255
            Index           =   23
            Left            =   360
            TabIndex        =   101
            Top             =   360
            Width           =   3855
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Set to:"
            Height          =   255
            Index           =   22
            Left            =   360
            TabIndex        =   100
            Top             =   840
            Width           =   1815
         End
      End
      Begin VB.Frame fraBranch 
         Caption         =   "Conditional Branch"
         Height          =   3735
         Left            =   120
         TabIndex        =   136
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.TextBox txtBranchItemAmount 
            Height          =   285
            Left            =   3480
            TabIndex        =   199
            Top             =   960
            Width           =   1815
         End
         Begin VB.HScrollBar scrlNegative 
            Height          =   255
            Left            =   120
            TabIndex        =   156
            Top             =   3360
            Value           =   1
            Width           =   5295
         End
         Begin VB.HScrollBar scrlPositive 
            Height          =   255
            Left            =   120
            TabIndex        =   154
            Top             =   2760
            Value           =   1
            Width           =   5295
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Player Variable"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   151
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Player Switch"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   150
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Has Item"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   149
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Player is Donator"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   148
            Top             =   1320
            Width           =   5175
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Knows Skill"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   147
            Top             =   1680
            Width           =   1215
         End
         Begin VB.ComboBox cmbBranchVar 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":0135
            Left            =   1560
            List            =   "frmEditor_Event.frx":0137
            TabIndex        =   146
            Text            =   "cmbBranchVar"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtBranchVarReq 
            Height          =   285
            Left            =   4320
            TabIndex        =   145
            Text            =   "0"
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox cmbVarReqOperator 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":0139
            Left            =   3480
            List            =   "frmEditor_Event.frx":014F
            Style           =   2  'Dropdown List
            TabIndex        =   144
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Level is"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   143
            Top             =   2040
            Width           =   975
         End
         Begin VB.ComboBox cmbLevelReqOperator 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":01B5
            Left            =   1560
            List            =   "frmEditor_Event.frx":01B7
            TabIndex        =   142
            Text            =   "cmbLevelReqOperator"
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox txtBranchLevelReq 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3360
            TabIndex        =   141
            Text            =   "0"
            Top             =   2040
            Width           =   855
         End
         Begin VB.ComboBox cmbBranchSwitch 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":01B9
            Left            =   1560
            List            =   "frmEditor_Event.frx":01BB
            TabIndex        =   140
            Text            =   "cmbBranchSwitch"
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox cmbBranchSwitchReq 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":01BD
            Left            =   3480
            List            =   "frmEditor_Event.frx":01C7
            TabIndex        =   139
            Text            =   "cmbBranchSwitchReq"
            Top             =   600
            Width           =   1815
         End
         Begin VB.ComboBox cmbBranchItem 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":01D8
            Left            =   1560
            List            =   "frmEditor_Event.frx":01DA
            TabIndex        =   138
            Text            =   "cmbBranchItem"
            Top             =   960
            Width           =   1695
         End
         Begin VB.ComboBox cmbBranchSkill 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":01DC
            Left            =   1560
            List            =   "frmEditor_Event.frx":01DE
            TabIndex        =   137
            Text            =   "cmbBranchSkill"
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label lblNegative 
            Caption         =   "Negative: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   157
            Top             =   3120
            Width           =   5295
         End
         Begin VB.Label lblPositive 
            Caption         =   "Positive: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   155
            Top             =   2520
            Width           =   5295
         End
         Begin VB.Label lblRandomLabel 
            Alignment       =   2  'Center
            Caption         =   "is"
            Height          =   255
            Index           =   0
            Left            =   3240
            TabIndex        =   153
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblRandomLabel 
            Alignment       =   2  'Center
            Caption         =   "is"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   152
            Top             =   600
            Width           =   255
         End
      End
      Begin VB.Frame fraChangeSkill 
         Caption         =   "Change Player Skills"
         Height          =   3735
         Left            =   120
         TabIndex        =   159
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.ComboBox cmbChangeSkills 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":01E0
            Left            =   720
            List            =   "frmEditor_Event.frx":01E2
            Style           =   2  'Dropdown List
            TabIndex        =   162
            Top             =   360
            Width           =   4695
         End
         Begin VB.OptionButton optChangeSkills 
            Caption         =   "Teach"
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   161
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton optChangeSkills 
            Caption         =   "Remove"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   160
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Skill:"
            Height          =   255
            Index           =   28
            Left            =   120
            TabIndex        =   163
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame fraChangePK 
         Caption         =   "Set Player PK"
         Height          =   3735
         Left            =   120
         TabIndex        =   173
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.OptionButton optChangePK 
            Caption         =   "No"
            Height          =   255
            Index           =   0
            Left            =   2040
            TabIndex        =   175
            Top             =   1200
            Width           =   1455
         End
         Begin VB.OptionButton optChangePK 
            Caption         =   "Yes"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   174
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame fraSpawnNPC 
         Caption         =   "Spawn NPC"
         Height          =   3735
         Left            =   120
         TabIndex        =   201
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.ComboBox cmbSpawnNPC 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":01E4
            Left            =   120
            List            =   "frmEditor_Event.frx":01E6
            Style           =   2  'Dropdown List
            TabIndex        =   202
            Top             =   480
            Width           =   5295
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "NPC:"
            Height          =   255
            Index           =   42
            Left            =   120
            TabIndex        =   203
            Top             =   240
            Width           =   3735
         End
      End
   End
   Begin VB.CommandButton cmdSwitchesVariables 
      Caption         =   "Switch/Variable"
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Info"
      Height          =   7695
      Left            =   3360
      TabIndex        =   5
      Top             =   120
      Width           =   6015
      Begin VB.Frame fraRandom 
         Caption         =   "Conditions"
         Height          =   1575
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   3975
         Begin VB.TextBox txtPlayerVariable 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3240
            TabIndex        =   38
            Top             =   240
            Width           =   615
         End
         Begin VB.ComboBox cmbPlayerVar 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":01E8
            Left            =   1200
            List            =   "frmEditor_Event.frx":01EA
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkPlayerVar 
            Caption         =   "Variable"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox cmbPlayerSwitch 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":01EC
            Left            =   1200
            List            =   "frmEditor_Event.frx":01EE
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox chkPlayerSwitch 
            Caption         =   "Switch"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   975
         End
         Begin VB.ComboBox cmbHasItem 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":01F0
            Left            =   1200
            List            =   "frmEditor_Event.frx":01F2
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1200
            Width           =   2655
         End
         Begin VB.CheckBox chkHasItem 
            Caption         =   "Has Item"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1200
            Width           =   975
         End
         Begin VB.ComboBox cmbPlayerSwitchCompare 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":01F4
            Left            =   2640
            List            =   "frmEditor_Event.frx":01FE
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   720
            Width           =   1215
         End
         Begin VB.ComboBox cmbPlayerVarCompare 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":020F
            Left            =   2640
            List            =   "frmEditor_Event.frx":0211
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblRandomLabel 
            Alignment       =   2  'Center
            Caption         =   "is"
            Height          =   255
            Index           =   5
            Left            =   2400
            TabIndex        =   40
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblRandomLabel 
            Alignment       =   2  'Center
            Caption         =   "is"
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   39
            Top             =   795
            Width           =   255
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Commands"
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   6840
         Width           =   5775
         Begin VB.CommandButton cmdSubEventEdit 
            Caption         =   "Edit"
            Height          =   375
            Left            =   2280
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdSubEventUp 
            Caption         =   "/\"
            Height          =   375
            Left            =   5160
            TabIndex        =   13
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdSubEventDown 
            Caption         =   "\/"
            Height          =   375
            Left            =   4560
            TabIndex        =   12
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdSubEventRemove 
            Caption         =   "Remove"
            Height          =   375
            Left            =   1200
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdSubEventAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   3255
      End
      Begin VB.ListBox lstSubEvents 
         Height          =   4545
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   5775
      End
      Begin VB.Frame fraGraphic 
         Caption         =   "Graphic"
         Height          =   2055
         Left            =   4200
         TabIndex        =   216
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
         Begin VB.CommandButton cmdGraphicClose 
            Caption         =   "Close"
            Height          =   255
            Left            =   120
            TabIndex        =   222
            Top             =   1680
            Width           =   1455
         End
         Begin VB.PictureBox picGraphic 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   960
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   221
            Top             =   240
            Width           =   615
         End
         Begin VB.HScrollBar scrlGraphic 
            Height          =   255
            Left            =   120
            TabIndex        =   219
            Top             =   1320
            Width           =   1455
         End
         Begin VB.HScrollBar scrlCurGraphic 
            Height          =   255
            Left            =   120
            Max             =   2
            TabIndex        =   218
            Top             =   720
            Width           =   735
         End
         Begin VB.CheckBox chkAnimated 
            Caption         =   "Animated?"
            Height          =   375
            Left            =   120
            TabIndex        =   217
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblGraphic 
            Caption         =   "Sprite#0: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   220
            Top             =   1080
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Map Tile only"
         Height          =   2055
         Left            =   4200
         TabIndex        =   41
         Top             =   120
         Width           =   1695
         Begin VB.CommandButton cmbGraphic 
            Caption         =   "Graphic"
            Height          =   255
            Left            =   120
            TabIndex        =   215
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CheckBox chkWalkthrought 
            Caption         =   "Walk through?"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbTrigger 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":0213
            Left            =   120
            List            =   "frmEditor_Event.frx":021D
            TabIndex        =   42
            Text            =   "cmbTrigger"
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Event List"
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   200
         Text            =   "Search..."
         Top             =   240
         Width           =   2895
      End
      Begin VB.ListBox lstIndex 
         Height          =   6885
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Events"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ListIndex As Long
Private Sub cmbBranchItem_Click()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(2) = cmbBranchItem.ListIndex + 1
End Sub

Private Sub cmbBranchSkill_Click()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(2) = cmbBranchSkill.ListIndex + 1
End Sub

Private Sub cmbBranchSwitch_Click()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(5) = cmbBranchSwitch.ListIndex
End Sub

Private Sub cmbBranchSwitchReq_Click()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(2) = cmbBranchSwitchReq.ListIndex
End Sub

Private Sub cmbBranchVar_Click()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(6) = cmbBranchVar.ListIndex
End Sub

Private Sub cmbGraphic_Click()
    fraGraphic.visible = True
End Sub

Private Sub cmbHasItem_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).HasItemIndex = cmbHasItem.ListIndex + 1
End Sub

Private Sub cmbChangeSkills_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(1) = cmbChangeSkills.ListIndex + 1
End Sub

Private Sub cmbChatBubbleTarget_click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(2) = cmbChatBubbleTarget.ListIndex + 1
End Sub

Private Sub cmbLabel_Ok_Click()
    fraLabeling.visible = False
    SendSwitchesAndVariables
End Sub

Private Sub cmbLevelReqOperator_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(5) = cmbLevelReqOperator.ListIndex
End Sub

Private Sub cmbPlayerSwitch_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).SwitchIndex = cmbPlayerSwitch.ListIndex
End Sub

Private Sub cmbPlayerSwitchCompare_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).SwitchCompare = cmbPlayerSwitchCompare.ListIndex
End Sub

Private Sub cmbPlayerSwitchSet_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(2) = cmbPlayerSwitchSet.ListIndex
End Sub

Private Sub cmbPlayerVar_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).VariableIndex = cmbPlayerVar.ListIndex
End Sub

Private Sub cmbPlayerVarCompare_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).VariableCompare = cmbPlayerVarCompare.ListIndex
End Sub

Private Sub cmbSetAccess_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(1) = cmbSetAccess.ListIndex
End Sub

Private Sub cmbSwitch_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(1) = cmbSwitch.ListIndex
End Sub

Private Sub cmbTrigger_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).Trigger = cmbTrigger.ListIndex
End Sub

Private Sub cmbVariable_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(1) = cmbVariable.ListIndex
End Sub

Private Sub cmbVarReqOperator_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(5) = cmbVarReqOperator.ListIndex
End Sub

Private Sub cmdAddMenuOption_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    With Events(EditorIndex).SubEvents(ListIndex)
        ReDim Preserve .data(1 To UBound(.data) + 1)
        ReDim Preserve .Text(1 To UBound(.data) + 1)
        .data(UBound(.data)) = 1
    End With
    lstMenuOptions.AddItem ": " & 1
End Sub

Private Sub cmdAddOk_Click()
    fraCommands.visible = False
End Sub

Private Sub cmdCommand_Click(Index As Integer)
    Dim Count As Long
    If Not (Events(EditorIndex).HasSubEvents) Then
        ReDim Events(EditorIndex).SubEvents(1 To 1)
        Events(EditorIndex).HasSubEvents = True
    Else
        Count = UBound(Events(EditorIndex).SubEvents) + 1
        ReDim Preserve Events(EditorIndex).SubEvents(1 To Count)
    End If
    Call Events_SetSubEventType(EditorIndex, UBound(Events(EditorIndex).SubEvents), Index)
    Call PopulateSubEventList
    fraCommands.visible = False
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex <= 0 Or EditorIndex > MAX_EVENTS Then Exit Sub
    ListIndex = 0
    ClearEvent EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Trim$(Events(EditorIndex).name), EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    Event_Changed(EditorIndex) = True
    EventEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Events", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdEditOk_Click()
    Call PopulateSubEventList
    fraEditCommand.visible = False
End Sub

Private Sub cmdGraphicClose_Click()
    fraGraphic.visible = False
End Sub

Private Sub cmdLabel_Cancel_Click()
    fraLabeling.visible = False
    RequestSwitchesAndVariables
End Sub

Private Sub cmdModifyMenuOption_Click()
    Dim optIdx As Long
    optIdx = lstMenuOptions.ListIndex + 1
    If optIdx < 1 Or optIdx > UBound(Events(EditorIndex).SubEvents(ListIndex).data) Then Exit Sub
    
    Events(EditorIndex).SubEvents(ListIndex).Text(optIdx + 1) = Trim$(txtMenuOptText.Text)
    Events(EditorIndex).SubEvents(ListIndex).data(optIdx) = scrlMenuOptDest.Value
    lstMenuOptions.List(optIdx - 1) = Trim$(txtMenuOptText.Text) & ": " & scrlMenuOptDest.Value
End Sub

Private Sub cmdRemoveMenuOption_Click()
    Dim Index As Long, I As Long
    
    Index = lstMenuOptions.ListIndex + 1
    If Index > 0 And Index < lstMenuOptions.ListCount And lstMenuOptions.ListCount > 0 Then
        For I = Index + 1 To lstMenuOptions.ListCount
            Events(EditorIndex).SubEvents(ListIndex).data(I - 1) = Events(EditorIndex).SubEvents(ListIndex).data(I)
            Events(EditorIndex).SubEvents(ListIndex).Text(I) = Events(EditorIndex).SubEvents(ListIndex).Text(I + 1)
        Next I
        ReDim Preserve Events(EditorIndex).SubEvents(ListIndex).data(1 To UBound(Events(EditorIndex).SubEvents(ListIndex).data) - 1)
        ReDim Preserve Events(EditorIndex).SubEvents(ListIndex).Text(1 To UBound(Events(EditorIndex).SubEvents(ListIndex).Text) - 1)
        Call PopulateSubEventConfig
    End If
End Sub

Private Sub cmdRename_Cancel_Click()
    Dim I As Long
    fraRenaming.visible = False
    RenameType = 0
    RenameIndex = 0
    lstSwitches.Clear
    For I = 1 To MAX_SWITCHES
        lstSwitches.AddItem CStr(I) & ". " & Trim$(Switches(I))
    Next
    lstSwitches.ListIndex = 0
    
    lstVariables.Clear
    For I = 1 To MAX_VARIABLES
        lstVariables.AddItem CStr(I) & ". " & Trim$(Variables(I))
    Next
    lstVariables.ListIndex = 0
End Sub

Private Sub cmdRename_Ok_Click()
    Dim I As Long
    Select Case RenameType
        Case 1
            'Variable
            If RenameIndex > 0 And RenameIndex <= MAX_VARIABLES + 1 Then
                Variables(RenameIndex) = txtRename.Text
                fraRenaming.visible = False
                RenameType = 0
                RenameIndex = 0
            End If
        Case 2
            'Switch
            If RenameIndex > 0 And RenameIndex <= MAX_SWITCHES + 1 Then
                Switches(RenameIndex) = txtRename.Text
                fraRenaming.visible = False
                RenameType = 0
                RenameIndex = 0
            End If
    End Select
    
    lstSwitches.Clear
    For I = 1 To MAX_SWITCHES
        lstSwitches.AddItem CStr(I) & ". " & Trim$(Switches(I))
    Next
    lstSwitches.ListIndex = 0
    
    lstVariables.Clear
    For I = 1 To MAX_VARIABLES
        lstVariables.AddItem CStr(I) & ". " & Trim$(Variables(I))
    Next
    lstVariables.ListIndex = 0
End Sub

Private Sub cmdRenameSwitch_Click()
    If lstSwitches.ListIndex > -1 And lstSwitches.ListIndex < MAX_SWITCHES Then
        fraRenaming.visible = True
        lblEditing.Caption = "Editing Switch #" & CStr(lstSwitches.ListIndex + 1)
        txtRename.Text = Switches(lstSwitches.ListIndex + 1)
        RenameType = 2
        RenameIndex = lstSwitches.ListIndex + 1
    End If
End Sub

Private Sub cmdRenameVariable_Click()
    If lstVariables.ListIndex > -1 And lstVariables.ListIndex < MAX_VARIABLES Then
        fraRenaming.visible = True
        lblEditing.Caption = "Editing Variable #" & CStr(lstVariables.ListIndex + 1)
        txtRename.Text = Variables(lstVariables.ListIndex + 1)
        RenameType = 1
        RenameIndex = lstVariables.ListIndex + 1
    End If
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call EventEditorOk
    ListIndex = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Events", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call EventEditorCancel
    ListIndex = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Events", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub cmdSubEventAdd_Click()
    fraCommands.visible = True
End Sub

Private Sub cmdSubEventDown_Click()
    Dim Index As Long
    Index = lstSubEvents.ListIndex + 1
    If Index > 0 And Index < lstSubEvents.ListCount Then
        Dim temp As SubEventRec
        temp = Events(EditorIndex).SubEvents(Index)
        Events(EditorIndex).SubEvents(Index) = Events(EditorIndex).SubEvents(Index + 1)
        Events(EditorIndex).SubEvents(Index + 1) = temp
        Call PopulateSubEventList
    End If
End Sub

Private Sub cmdSubEventEdit_Click()
    If ListIndex > 0 And ListIndex <= lstSubEvents.ListCount Then
        fraEditCommand.visible = True
        PopulateSubEventConfig
    End If
End Sub

Private Sub cmdSubEventRemove_Click()
    Dim I As Long
    If ListIndex > 0 And ListIndex <= lstSubEvents.ListCount Then
        For I = ListIndex + 1 To lstSubEvents.ListCount
            Events(EditorIndex).SubEvents(I - 1) = Events(EditorIndex).SubEvents(I)
        Next I
        If lstSubEvents.ListCount = 1 Then
            Events(EditorIndex).HasSubEvents = False
            Erase Events(EditorIndex).SubEvents
        Else
            ReDim Preserve Events(EditorIndex).SubEvents(1 To lstSubEvents.ListCount - 1)
        End If
        Call PopulateSubEventList
    End If
End Sub

Private Sub cmdSubEventUp_Click()
    Dim Index As Long
    Index = lstSubEvents.ListIndex + 1
    If Index > 1 And Index <= lstSubEvents.ListCount Then
        Dim temp As SubEventRec
        temp = Events(EditorIndex).SubEvents(Index)
        Events(EditorIndex).SubEvents(Index) = Events(EditorIndex).SubEvents(Index - 1)
        Events(EditorIndex).SubEvents(Index - 1) = temp
        Call PopulateSubEventList
    End If
End Sub

Private Sub cmdSwitchesVariables_Click()
Dim I As Long
    fraLabeling.visible = True
    lstSwitches.Clear
    For I = 1 To MAX_SWITCHES
        lstSwitches.AddItem CStr(I) & ". " & Trim$(Switches(I))
    Next
    lstSwitches.ListIndex = 0
    
    lstVariables.Clear
    For I = 1 To MAX_VARIABLES
        lstVariables.AddItem CStr(I) & ". " & Trim$(Variables(I))
    Next
    lstVariables.ListIndex = 0
End Sub

Private Sub Form_Load()
    Dim I As Long
    'Move windows to right places
    frmEditor_Events.width = 9600
    frmEditor_Events.height = 8835
    fraEditCommand.Left = 232
    fraEditCommand.Top = 152
    fraCommands.Left = 232
    fraCommands.Top = 152
    fraLabeling.width = 609
    fraLabeling.height = 513
    
    ListIndex = 0

    scrlOpenShop.Max = MAX_SHOPS
    scrlGiveItemID.Max = MAX_ITEMS
    scrlPlayAnimationAnim.Max = MAX_ANIMATIONS
    scrlWarpMap.Max = MAX_MAPS
    scrlGraphic.Max = Count_Event
    
    cmbLevelReqOperator.Clear
    cmbPlayerVarCompare.Clear
    cmbVarReqOperator.Clear
    For I = 0 To ComparisonOperator_Count - 1
        cmbLevelReqOperator.AddItem GetComparisonOperatorName(I)
        cmbPlayerVarCompare.AddItem GetComparisonOperatorName(I)
        cmbVarReqOperator.AddItem GetComparisonOperatorName(I)
    Next
    
    cmbHasItem.Clear
    cmbBranchItem.Clear
    For I = 1 To MAX_ITEMS
        cmbHasItem.AddItem Trim$(Item(I).name)
        cmbBranchItem.AddItem Trim$(Item(I).name)
    Next
    
    cmbSwitch.Clear
    cmbPlayerSwitch.Clear
    cmbBranchSwitch.Clear
    For I = 1 To MAX_SWITCHES
        cmbSwitch.AddItem I & ". " & Switches(I)
        cmbPlayerSwitch.AddItem I & ". " & Switches(I)
        cmbBranchSwitch.AddItem I & ". " & Switches(I)
    Next
    
    cmbVariable.Clear
    cmbPlayerVar.Clear
    cmbBranchVar.Clear
    For I = 1 To MAX_VARIABLES
        cmbVariable.AddItem I & ". " & Variables(I)
        cmbPlayerVar.AddItem I & ". " & Variables(I)
        cmbBranchVar.AddItem I & ". " & Variables(I)
    Next
    
    cmbBranchSkill.Clear
    cmbChangeSkills.Clear
    For I = 1 To MAX_SPELLS
        cmbBranchSkill.AddItem Trim$(spell(I).name)
        cmbChangeSkills.AddItem Trim$(spell(I).name)
    Next
    
    cmbChatBubbleTarget.Clear
    For I = 1 To MAX_MAP_NPCS
        If map.NPC(I) <= 0 Then
            cmbChatBubbleTarget.AddItem CStr(I) & ". "
        Else
            cmbChatBubbleTarget.AddItem CStr(I) & ". " & Trim$(NPC(map.NPC(I)).name)
        End If
    Next
End Sub

Private Sub chkHasItem_Click()
    If chkHasItem.Value = 0 Then cmbHasItem.Enabled = False Else cmbHasItem.Enabled = True
    Events(EditorIndex).chkHasItem = chkHasItem.Value
End Sub

Private Sub chkPlayerSwitch_Click()
    If chkPlayerSwitch.Value = 0 Then
        cmbPlayerSwitch.Enabled = False
        cmbPlayerSwitchCompare.Enabled = False
    Else
        cmbPlayerSwitch.Enabled = True
        cmbPlayerSwitchCompare.Enabled = True
    End If
    Events(EditorIndex).chkSwitch = chkPlayerSwitch.Value
End Sub

Private Sub chkPlayerVar_Click()
    If chkPlayerVar.Value = 0 Then
        cmbPlayerVar.Enabled = False
        txtPlayerVariable.Enabled = False
        cmbPlayerVarCompare.Enabled = False
    Else
        cmbPlayerVar.Enabled = True
        txtPlayerVariable.Enabled = True
        cmbPlayerVarCompare.Enabled = True
    End If
    Events(EditorIndex).chkVariable = chkPlayerVar.Value
End Sub

Private Sub chkWalkthrought_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).WalkThrought = chkWalkthrought.Value
End Sub
Private Sub chkAnimated_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).Animated = chkAnimated.Value
End Sub

Private Sub scrlGraphic_Change()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).Graphic(scrlCurGraphic.Value) = scrlGraphic.Value
    lblGraphic.Caption = "Sprite#" & scrlCurGraphic.Value & ": " & scrlGraphic.Value
End Sub
Private Sub scrlCurGraphic_Change()
If EditorIndex = 0 Then Exit Sub
    scrlGraphic.Value = Events(EditorIndex).Graphic(scrlCurGraphic.Value)
    lblGraphic.Caption = "Sprite#" & scrlCurGraphic.Value & ": " & scrlGraphic.Value
End Sub
Private Sub lstIndex_Click()
    EventEditorInit
End Sub

Private Sub lstMenuOptions_Click()
    Dim optIdx As Long
    optIdx = lstMenuOptions.ListIndex + 1
    If optIdx < 1 Or optIdx > UBound(Events(EditorIndex).SubEvents(ListIndex).data) Then Exit Sub
    
    txtMenuOptText.Text = Trim$(Events(EditorIndex).SubEvents(ListIndex).Text(optIdx + 1))
    If Events(EditorIndex).SubEvents(ListIndex).data(optIdx) <= 0 Then Events(EditorIndex).SubEvents(ListIndex).data(optIdx) = 1
    scrlMenuOptDest.Value = Events(EditorIndex).SubEvents(ListIndex).data(optIdx)
End Sub

Private Sub lstSubEvents_Click()
    ListIndex = lstSubEvents.ListIndex + 1
    If ListIndex > 0 And ListIndex < lstSubEvents.ListCount Then
        cmdSubEventDown.Enabled = True
    Else
        cmdSubEventDown.Enabled = False
    End If
    If ListIndex > 1 And ListIndex <= lstSubEvents.ListCount Then
        cmdSubEventUp.Enabled = True
    Else
        cmdSubEventUp.Enabled = False
    End If
End Sub

Private Sub lstSubEvents_DblClick()
    If ListIndex > 0 And ListIndex <= lstSubEvents.ListCount Then
        fraEditCommand.visible = True
        PopulateSubEventConfig
    End If
End Sub

Private Sub optCondition_Index_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub

    cmbBranchVar.Enabled = False
    cmbVarReqOperator.Enabled = False
    txtBranchVarReq.Enabled = False
    cmbBranchSwitch.Enabled = False
    cmbBranchSwitchReq.Enabled = False
    cmbBranchItem.Enabled = False
    txtBranchItemAmount.Enabled = False
    cmbBranchSkill.Enabled = False
    cmbLevelReqOperator.Enabled = False
    txtBranchLevelReq.Enabled = False
    
    Select Case Index
        Case 0
            cmbBranchVar.Enabled = True
            cmbVarReqOperator.Enabled = True
            txtBranchVarReq.Enabled = True
        Case 1
            cmbBranchSwitch.Enabled = True
            cmbBranchSwitchReq.Enabled = True
        Case 2
            cmbBranchItem.Enabled = True
            txtBranchItemAmount.Enabled = True
        Case 4
            cmbBranchSkill.Enabled = True
        Case 5
            cmbLevelReqOperator.Enabled = True
            txtBranchLevelReq.Enabled = True
    End Select

    Events(EditorIndex).SubEvents(ListIndex).data(1) = Index
End Sub

Private Sub optExpAction_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(2) = Index
End Sub

Private Sub optChangePK_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(1) = Index
End Sub

Private Sub optChangeSkills_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(2) = Index
End Sub

Private Sub optChannel_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(2) = Index
End Sub

Private Sub optChatBubbleTarget_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    If Index = 0 Then
        cmbChatBubbleTarget.visible = False
    ElseIf Index = 1 Then
        cmbChatBubbleTarget.visible = True
    End If
    Events(EditorIndex).SubEvents(ListIndex).data(1) = Index
End Sub

Private Sub optItemOperation_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(3) = Index
End Sub

Private Sub optLevelAction_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(2) = Index
End Sub

Private Sub optOpenEventType_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(3) = Index
End Sub

Private Sub optVariableAction_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(2) = Index
    Select Case Index
        Case 0, 1, 2
            txtVariableData(0).Enabled = True
            txtVariableData(1).Enabled = False
            txtVariableData(2).Enabled = False
        Case 3
            txtVariableData(0).Enabled = False
            txtVariableData(1).Enabled = True
            txtVariableData(2).Enabled = True
    End Select
End Sub

Private Sub scrlAddText_Colour_Change()
    frmEditor_Events.lblAddText_Colour.Caption = "Color: " & GetColorString(frmEditor_Events.scrlAddText_Colour.Value)
    Events(EditorIndex).SubEvents(ListIndex).data(1) = scrlAddText_Colour.Value
End Sub

Private Sub scrlCustomScript_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblCustomScript.Caption = "Case: " & scrlCustomScript.Value
    Events(EditorIndex).SubEvents(ListIndex).data(1) = scrlCustomScript.Value
End Sub

Private Sub scrlChangeExp_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblChangeExp.Caption = "Exp: " & scrlChangeExp.Value
    Events(EditorIndex).SubEvents(ListIndex).data(1) = scrlChangeExp.Value
End Sub

Private Sub scrlGiveItemAmount_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblGiveItemAmount.Caption = "Amount: " & scrlGiveItemAmount.Value
    Events(EditorIndex).SubEvents(ListIndex).data(2) = scrlGiveItemAmount.Value
End Sub

Private Sub scrlGiveItemID_Change()
    lblGiveItemID.Caption = "Item: " & scrlGiveItemID.Value & "-" & Item(scrlGiveItemID.Value).name
    Events(EditorIndex).SubEvents(ListIndex).data(1) = scrlGiveItemID.Value
End Sub

Private Sub scrlGOTO_Change()
    lblGOTO.Caption = "Goto: " & scrlGOTO.Value
    Events(EditorIndex).SubEvents(ListIndex).data(1) = scrlGOTO.Value
End Sub

Private Sub scrlChangeLevel_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblChangeLevel.Caption = "Level: " & scrlChangeLevel.Value
    Events(EditorIndex).SubEvents(ListIndex).data(1) = scrlChangeLevel.Value
End Sub

Private Sub scrlMenuOptDest_Change()
    lblMenuOptDest.Caption = "Destination: " & scrlMenuOptDest.Value
End Sub

Private Sub scrlOpenEventX_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblOpenEventX.Caption = "X: " & scrlOpenEventX.Value
    Events(EditorIndex).SubEvents(ListIndex).data(1) = scrlOpenEventX.Value
End Sub
Private Sub scrlOpenEventY_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblOpenEventY.Caption = "Y: " & scrlOpenEventY.Value
    Events(EditorIndex).SubEvents(ListIndex).data(2) = scrlOpenEventY.Value
End Sub

Private Sub scrlOpenShop_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblOpenShop.Caption = "Open Shop: " & scrlOpenShop.Value & "-" & Shop(scrlOpenShop.Value).name
    Events(EditorIndex).SubEvents(ListIndex).data(1) = scrlOpenShop.Value
End Sub

Private Sub scrlPlayAnimationAnim_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblPlayAnimationAnim.Caption = "Animation: " & scrlPlayAnimationAnim.Value & "-" & Animation(scrlPlayAnimationAnim.Value).name
    Events(EditorIndex).SubEvents(ListIndex).data(1) = scrlPlayAnimationAnim.Value
End Sub

Private Sub scrlPlayAnimationX_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    If scrlPlayAnimationX.Value >= 0 Then
        lblPlayAnimationX.Caption = "X: " & scrlPlayAnimationX.Value
    Else
        lblPlayAnimationX.Caption = "X: Player's X Position"
    End If
    Events(EditorIndex).SubEvents(ListIndex).data(2) = scrlPlayAnimationX.Value
End Sub

Private Sub scrlPlayAnimationY_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    If scrlPlayAnimationY.Value >= 0 Then
        lblPlayAnimationY.Caption = "Y: " & scrlPlayAnimationY.Value
    Else
        lblPlayAnimationY.Caption = "Y: Player's Y Position"
    End If
    Events(EditorIndex).SubEvents(ListIndex).data(3) = scrlPlayAnimationY.Value
End Sub

Private Sub scrlPositive_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblPositive.Caption = "Positive: " & scrlPositive.Value
    Events(EditorIndex).SubEvents(ListIndex).data(3) = scrlPositive.Value
End Sub
Private Sub scrlNegative_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblNegative.Caption = "Negative: " & scrlNegative.Value
    Events(EditorIndex).SubEvents(ListIndex).data(4) = scrlNegative.Value
End Sub

Private Sub scrlWarpMap_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblWarpMap.Caption = "Map: " & scrlWarpMap.Value
    Events(EditorIndex).SubEvents(ListIndex).data(1) = scrlWarpMap.Value
End Sub

Private Sub scrlWarpX_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblWarpX.Caption = "X: " & scrlWarpX.Value
    Events(EditorIndex).SubEvents(ListIndex).data(2) = scrlWarpX.Value
End Sub

Private Sub scrlWarpY_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblWarpY.Caption = "Y: " & scrlWarpY.Value
    Events(EditorIndex).SubEvents(ListIndex).data(3) = scrlWarpY.Value
End Sub

Private Sub txtAddText_Text_Change()
    Events(EditorIndex).SubEvents(ListIndex).Text(1) = Trim$(txtAddText_Text.Text)
End Sub

Private Sub txtBranchItemAmount_Change()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(5) = Val(txtBranchItemAmount.Text)
End Sub

Private Sub txtBranchLevelReq_Change()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(2) = Val(txtBranchLevelReq.Text)
End Sub

Private Sub txtBranchVarReq_Change()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(2) = Val(txtBranchVarReq.Text)
End Sub

Private Sub txtChatbubbleText_Change()
    Events(EditorIndex).SubEvents(ListIndex).Text(1) = txtChatbubbleText.Text
End Sub

Private Sub txtMenuQuery_Change()
    Events(EditorIndex).SubEvents(ListIndex).Text(1) = txtMenuQuery.Text
End Sub
Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_EVENTS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Events(EditorIndex).name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Events(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PopulateSubEventList()
    Dim tempIndex As Long, I As Long
    tempIndex = lstSubEvents.ListIndex
    
    lstSubEvents.Clear
    If Events(EditorIndex).HasSubEvents Then
        For I = 1 To UBound(Events(EditorIndex).SubEvents)
            lstSubEvents.AddItem I & ": " & GetEventTypeName(EditorIndex, I)
        Next I
    End If
    cmdSubEventRemove.Enabled = Events(EditorIndex).HasSubEvents
    
    If tempIndex >= 0 And tempIndex < lstSubEvents.ListCount - 1 Then lstSubEvents.ListIndex = tempIndex
    Call PopulateSubEventConfig
End Sub

Public Sub PopulateSubEventConfig()
    Dim I As Long
    If Not (fraEditCommand.visible) Then Exit Sub
    If ListIndex = 0 Then Exit Sub
    HideMenus
    'Ensure Capacity
    Call Events_SetSubEventType(EditorIndex, ListIndex, Events(EditorIndex).SubEvents(ListIndex).Type)
    
    With Events(EditorIndex).SubEvents(ListIndex)
        Select Case .Type
            Case Evt_Message
                txtPlayerText.Text = Trim$(.Text(1))
                fraPlayerText.visible = True
            Case Evt_Menu
                txtMenuQuery.Text = Trim$(.Text(1))
                lstMenuOptions.Clear
                For I = 2 To UBound(.Text)
                    lstMenuOptions.AddItem Trim$(.Text(I)) & ": " & .data(I - 1)
                Next I
                scrlMenuOptDest.Max = UBound(Events(EditorIndex).SubEvents)
                fraMenu.visible = True
            Case Evt_OpenShop
                If .data(1) < 1 Or .data(1) > MAX_SHOPS Then .data(1) = 1
                
                scrlOpenShop.Value = .data(1)
                Call scrlOpenShop_Change
                fraOpenShop.visible = True
            Case Evt_GiveItem
                If .data(1) < 1 Or .data(1) > MAX_ITEMS Then .data(1) = 1
                If .data(2) < 1 Then .data(2) = 1
                optItemOperation(.data(3)).Value = True
                scrlGiveItemID.Value = .data(1)
                scrlGiveItemAmount.Value = .data(2)
                Call scrlGiveItemID_Change
                Call scrlGiveItemAmount_Change
                fraGiveItem.visible = True
            Case Evt_PlayAnimation
                If .data(1) < 1 Or .data(1) > MAX_ANIMATIONS Then .data(1) = 1
                
                scrlPlayAnimationAnim.Value = .data(1)
                scrlPlayAnimationX.Value = .data(2)
                scrlPlayAnimationY.Value = .data(3)
                Call scrlPlayAnimationAnim_Change
                Call scrlPlayAnimationX_Change
                Call scrlPlayAnimationY_Change
                fraAnimation.visible = True
            Case Evt_Warp
                If .data(1) < 1 Or .data(1) > MAX_MAPS Then .data(1) = 1
                
                scrlWarpMap.Value = .data(1)
                scrlWarpX.Value = .data(2)
                scrlWarpY.Value = .data(3)
                Call scrlWarpMap_Change
                Call scrlWarpX_Change
                Call scrlWarpY_Change
                fraMapWarp.visible = True
            Case Evt_GOTO
                If .data(1) < 1 Or .data(1) > UBound(Events(EditorIndex).SubEvents) Then .data(1) = 1
                
                scrlGOTO.Max = UBound(Events(EditorIndex).SubEvents)
                scrlGOTO.Value = .data(1)
                Call scrlGOTO_Change
                fraGoTo.visible = True
            Case Evt_Switch
                cmbSwitch.ListIndex = .data(1)
                cmbPlayerSwitchSet.ListIndex = .data(2)
                fraChangeSwitch.visible = True
            Case Evt_Variable
                optVariableAction(.data(1)).Value = True
                If .data(1) = 3 Then
                    txtVariableData(1) = .data(2)
                    txtVariableData(2) = .data(3)
                Else
                    txtVariableData(0) = .data(2)
                End If
                fraChangeVariable.visible = True
            Case Evt_AddText
                txtAddText_Text.Text = Trim$(.Text(1))
                scrlAddText_Colour.Value = .data(1)
                optChannel(.data(2)).Value = True
                fraAddText.visible = True
            Case Evt_Chatbubble
                txtChatbubbleText.Text = Trim$(.Text(1))
                optChatBubbleTarget(.data(1)).Value = True
                cmbChatBubbleTarget.ListIndex = .data(2) - 1
                fraChatbubble.visible = True
            Case Evt_Branch
                scrlPositive.Max = UBound(Events(EditorIndex).SubEvents)
                scrlNegative.Max = UBound(Events(EditorIndex).SubEvents)
                scrlPositive.Value = .data(3)
                scrlNegative.Value = .data(4)
                optCondition_Index(.data(1)) = True
                Select Case .data(1)
                    Case 0
                        cmbBranchVar.ListIndex = .data(6)
                        txtBranchVarReq.Text = .data(2)
                        cmbVarReqOperator.ListIndex = .data(5)
                    Case 1
                        cmbBranchSwitch.ListIndex = .data(5)
                        cmbBranchSwitchReq.ListIndex = .data(2)
                    Case 2
                        cmbBranchItem.ListIndex = .data(2) - 1
                        txtBranchItemAmount.Text = .data(5)
                    Case 4
                        cmbBranchSkill.ListIndex = .data(2) - 1
                    Case 5
                        cmbLevelReqOperator.ListIndex = .data(5)
                        txtBranchLevelReq.Text = .data(2)
                End Select
                fraBranch.visible = True
            Case Evt_ChangeSkill
                cmbChangeSkills.ListIndex = .data(1) - 1
                optChangeSkills(.data(2)).Value = True
                fraChangeSkill.visible = True
            Case Evt_ChangeLevel
                scrlChangeLevel.Value = .data(1)
                optLevelAction(.data(2)).Value = True
                fraChangeLevel.visible = True
            Case Evt_ChangePK
                optChangePK(.data(1)).Value = True
                fraChangePK.visible = True
            Case Evt_ChangeExp
                scrlChangeExp.Value = .data(1)
                optExpAction(.data(2)).Value = True
                fraChangeExp.visible = True
            Case Evt_SetAccess
                cmbSetAccess.ListIndex = .data(1)
                fraSetAccess.visible = True
            Case Evt_CustomScript
                scrlCustomScript.Value = .data(1)
                fraCustomScript.visible = True
            Case Evt_OpenEvent
                scrlOpenEventX.Value = .data(1)
                scrlOpenEventY.Value = .data(2)
                optOpenEventType(.data(3)).Value = True
                cmbOpenEventType.ListIndex = .data(4)
                fraOpenEvent.visible = True
            Case Evt_SpawnNPC
                cmbSpawnNPC.ListIndex = .data(1) - 1
                fraSpawnNPC.visible = True
            Case Evt_Changegraphic
                scrlChangeGraphicX.Value = .data(1)
                scrlChangeGraphicY.Value = .data(2)
                scrlChangeGraphic.Value = .data(3)
                cmbChangeGraphicType.ListIndex = .data(4)
                fraChangeGraphic.visible = True
        End Select
    End With
End Sub
Private Sub HideMenus()
    fraPlayerText.visible = False
    fraMenu.visible = False
    fraOpenShop.visible = False
    fraGiveItem.visible = False
    fraAnimation.visible = False
    fraMapWarp.visible = False
    fraGoTo.visible = False
    fraChangeSwitch.visible = False
    fraChangeVariable.visible = False
    fraAddText.visible = False
    fraChatbubble.visible = False
    fraBranch.visible = False
    fraChangeLevel.visible = False
    fraChangeSkill.visible = False
    fraChangePK.visible = False
    fraSetAccess.visible = False
    fraCustomScript.visible = False
    fraOpenEvent.visible = False
    fraChangeExp.visible = False
    fraSpawnNPC.visible = False
    fraChangeGraphic.visible = False
End Sub

Private Sub txtPlayerText_Change()
    Events(EditorIndex).SubEvents(ListIndex).Text(1) = txtPlayerText.Text
End Sub

Private Sub txtPlayerVariable_Change()
    Events(EditorIndex).VariableCondition = Val(txtPlayerVariable.Text)
End Sub

Private Sub txtVariableData_Change(Index As Integer)
    Select Case Index
        Case 0: Events(EditorIndex).SubEvents(ListIndex).data(3) = Val(txtVariableData(0))
        Case 1: Events(EditorIndex).SubEvents(ListIndex).data(3) = Val(txtVariableData(1))
        Case 2: Events(EditorIndex).SubEvents(ListIndex).data(4) = Val(txtVariableData(2))
    End Select
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
Private Sub scrlChangeGraphicX_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblChangeGraphicX.Caption = "X: " & scrlChangeGraphicX.Value
    Events(EditorIndex).SubEvents(ListIndex).data(1) = scrlChangeGraphicX.Value
End Sub
Private Sub scrlChangeGraphicY_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblChangeGraphicY.Caption = "Y: " & scrlChangeGraphicY.Value
    Events(EditorIndex).SubEvents(ListIndex).data(2) = scrlChangeGraphicY.Value
End Sub
Private Sub scrlChangeGraphic_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblChangeGraphic.Caption = "Graphic#: " & scrlChangeGraphic.Value
    Events(EditorIndex).SubEvents(ListIndex).data(3) = scrlChangeGraphic.Value
End Sub

Private Sub cmbOpenEventType_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(4) = cmbOpenEventType.ListIndex
End Sub

Private Sub cmbChangeGraphicType_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).data(4) = cmbChangeGraphicType.ListIndex
End Sub
