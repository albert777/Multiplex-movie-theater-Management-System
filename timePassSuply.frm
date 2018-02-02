VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form timePassSuply 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   13590
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      Caption         =   "SAVE"
      Height          =   855
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Width           =   1935
   End
   Begin VB.ComboBox cmbNos 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "timePassSuply.frx":0000
      Left            =   2760
      List            =   "timePassSuply.frx":0022
      TabIndex        =   5
      Text            =   "1"
      Top             =   2520
      Width           =   2775
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1935
      Left            =   5760
      TabIndex        =   3
      Top             =   1560
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.TextBox txtTotal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox txtPrice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   2040
      Width           =   2775
   End
   Begin VB.ComboBox cmbItem 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2760
      TabIndex        =   0
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   ".00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   39
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11760
      TabIndex        =   8
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblStock 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   39
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9120
      TabIndex        =   4
      Top             =   3600
      Width           =   2535
   End
End
Attribute VB_Name = "timePassSuply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbItem_Click()
If rs.State Then rs.Close
rs.Open "select amount from food where id='" & getId(cmbItem.Text) & "'", con
txtPrice.Text = rs(0)
cmbNos.SetFocus
End Sub

Private Sub cmbNos_Click()
Dim stock As Integer
Dim nos As Integer
Dim amt As Double

nos = Val(cmbNos.Text)
amt = Val(txtPrice.Text)
txtTotal.Text = nos * amt

If rs.State Then rs.Close
rs.Open "select stock from food  where id='" & getId(cmbItem.Text) & "'", con
stock = Val(rs(0))

If (stock - Val(cmbNos.Text)) <= 0 Then
    If MsgBox("Only " & stock & " nos are available... Do you want to proceed ?", vbYesNo, "Less stock Information") = vbNo Then Exit Sub
End If

For i = 1 To MSHFlexGrid1.Rows - 1
    If getId(MSHFlexGrid1.TextMatrix(i, 1)) = getId(cmbItem.Text) Then Exit Sub
Next i
MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = MSHFlexGrid1.Rows - 1
MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 1) = cmbItem.Text
MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 2) = txtPrice.Text
MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 3) = cmbNos.Text
MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 4) = txtTotal.Text
lblTotal.Caption = Val(lblTotal.Caption) + Val(txtTotal.Text)
End Sub



Private Sub Command1_Click()
For i = 0 To MSHFlexGrid1.Rows - 1
    con.Execute "insert into sales values(getdate(),'" & getId(MSHFlexGrid1.TextMatrix(i, 1)) & "','" & MSHFlexGrid1.TextMatrix(i, 2) & "','" & MSHFlexGrid1.TextMatrix(i, 3) & "','" & MSHFlexGrid1.TextMatrix(i, 4) & "')"
    con.Execute "update set food set stock=stock-" & MSHFlexGrid1.TextMatrix(i, 3) & " where id='" & getId(MSHFlexGrid1.TextMatrix(i, 1)) & "'"
Next i
End Sub

Private Sub Form_Load()
connectdb
If rs.State Then rs.Close
Call bindComboBox(cmbItem, "select * from food ")
End Sub



Private Sub txtPrice_GotFocus()
Call cmbItem_Click
End Sub

Private Sub txtTotal_GotFocus()
Call cmbNos_Click
End Sub
 
