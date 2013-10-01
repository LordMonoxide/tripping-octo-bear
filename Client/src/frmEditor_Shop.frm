VERSION 5.00
Begin VB.Form frmEditor_Shop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shop Editor"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
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
   Icon            =   "frmEditor_Shop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5640
      TabIndex        =   19
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Shop Properties"
      Height          =   4455
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   5295
      Begin VB.TextBox txtCostValue2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         TabIndex        =   25
         Text            =   "1"
         Top             =   2040
         Width           =   615
      End
      Begin VB.ComboBox cmbCostItem2 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2040
         Width           =   3135
      End
      Begin VB.HScrollBar scrlShoptype 
         Height          =   255
         Left            =   2160
         Max             =   1
         TabIndex        =   22
         Top             =   600
         Width           =   3015
      End
      Begin VB.CommandButton cmdDeleteTrade 
         Caption         =   "Delete"
         Height          =   255
         Left            =   2760
         TabIndex        =   20
         Top             =   2400
         Width           =   2415
      End
      Begin VB.HScrollBar scrlBuy 
         Height          =   255
         Left            =   2160
         Max             =   1000
         Min             =   1
         TabIndex        =   18
         Top             =   960
         Value           =   100
         Width           =   3015
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   4455
      End
      Begin VB.ListBox lstTradeItem 
         Height          =   1500
         ItemData        =   "frmEditor_Shop.frx":3332
         Left            =   120
         List            =   "frmEditor_Shop.frx":334E
         TabIndex        =   10
         Top             =   2760
         Width           =   5055
      End
      Begin VB.ComboBox cmbItem 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtItemValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         TabIndex        =   8
         Text            =   "1"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cmbCostItem 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtCostValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         TabIndex        =   6
         Text            =   "1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   180
         Left            =   3960
         TabIndex        =   27
         Top             =   2040
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblTool 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   2040
         Width           =   450
      End
      Begin VB.Label lblShopType 
         AutoSize        =   -1  'True
         Caption         =   "Shop type: Shop"
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   1230
      End
      Begin VB.Label lblBuy 
         AutoSize        =   -1  'True
         Caption         =   "Buy Rate: 100%"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   180
         Left            =   3960
         TabIndex        =   14
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   180
         Left            =   3960
         TabIndex        =   12
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Shop List"
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   21
         Text            =   "Search..."
         Top             =   240
         Width           =   2895
      End
      Begin VB.ListBox lstIndex 
         Height          =   3660
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   4680
      Width           =   1455
   End
End
Attribute VB_Name = "frmEditor_Shop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()
    If LenB(Trim$(txtName)) = 0 Then
        Call MsgBox("Name required.")
    Else
        Call ShopEditorOk
    End If
End Sub

Private Sub cmdCancel_Click()
    Call ShopEditorCancel
End Sub

Private Sub cmdUpdate_Click()
Dim Index As Long
Dim tmpPos As Long

    tmpPos = lstTradeItem.ListIndex
    Index = lstTradeItem.ListIndex + 1
    If Index = 0 Then Exit Sub
    With Shop(EditorIndex).TradeItem(Index)
        .Item = cmbItem.ListIndex
        .ItemValue = Val(txtItemValue.Text)
        .CostItem = cmbCostItem.ListIndex
        .CostValue = Val(txtCostValue.Text)
        .CostItem2 = cmbCostItem2.ListIndex
        .CostValue2 = Val(txtCostValue2.Text)
    End With
    UpdateShopTrade tmpPos
End Sub

Private Sub cmdDeleteTrade_Click()
Dim Index As Long

    Index = lstTradeItem.ListIndex + 1
    If Index = 0 Then Exit Sub
    With Shop(EditorIndex).TradeItem(Index)
        .Item = 0
        .ItemValue = 0
        .CostItem = 0
        .CostValue = 0
    End With
    Call UpdateShopTrade
End Sub

Private Sub lstIndex_Click()
    ShopEditorInit
End Sub

Private Sub scrlBuy_Change()
    lblBuy.Caption = "Buy Rate: " & scrlBuy.Value & "%"
    Shop(EditorIndex).BuyRate = scrlBuy.Value
End Sub

Private Sub scrlShoptype_Change()
    Select Case scrlShoptype.Value
        Case 0
            lblShopType.Caption = "Shop Type: Shop"
        Case 1
            lblShopType.Caption = "Shop Type: Anvil"
        Case Else
            lblShopType.Caption = "Shop Type: None"
    End Select
    
    Shop(EditorIndex).ShopType = scrlShoptype.Value
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    tmpIndex = lstIndex.ListIndex
    Shop(EditorIndex).name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Shop(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
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
