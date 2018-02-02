VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form stock 
   Caption         =   "Form1"
   ClientHeight    =   7680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   13065
   Begin VB.Frame Frame2 
      Height          =   4425
      Left            =   7080
      TabIndex        =   8
      Top             =   2160
      Width           =   6975
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4215
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   7435
         _Version        =   393216
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
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4425
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   6735
      Begin VB.TextBox stocktxt 
         Height          =   495
         Left            =   1920
         TabIndex        =   10
         Top             =   2400
         Width           =   4335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ADD / UPDATE"
         Height          =   495
         Left            =   1920
         TabIndex        =   3
         Top             =   3240
         Width           =   2415
      End
      Begin VB.TextBox amounttxt 
         Height          =   495
         Left            =   1920
         TabIndex        =   2
         Top             =   1680
         Width           =   4335
      End
      Begin VB.TextBox nametxt 
         Height          =   495
         Left            =   1920
         TabIndex        =   1
         Top             =   960
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "Amount"
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Stock"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lblId 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim str, msg As String
If rs.State Then rs.Close
rs.Open "select count(*) from food where id='" & lblId.Caption & "'", con
If Val(rs(0)) <= 0 Then
    str = "insert into food values('" & nametxt.Text & "','" & amounttxt.Text & "','" & stocktxt.Text & "')"
Else
    If rs.State Then rs.Close
    rs.Open "select stock from food where id='" & lblId.Caption & "'", con
    If Val(rs(0)) <= 0 Then
        str = "update food set stock=" & Val(stocktxt.Text) & ",name='" & nametxt.Text & "', amount='" & Val(amounttxt.Text) & "' where id='" & lblId.Caption & "'"
        msg = "Stock Initiliased and updated"
    ElseIf Val(rs(0)) > 0 Then
        str = "update food set stock=" & Val(rs(0)) + Val(stocktxt.Text) & ",name='" & nametxt.Text & "', amount='" & Val(amounttxt.Text) & "' where id='" & lblId.Caption & "'"
        msg = "stock updated"
    End If
End If
con.Execute str
MsgBox msg
Unload Me
stock.Show
End Sub

Private Sub Command2_Click()
con.Execute "update food set stock=stock+  '" & Val(stocktxt.Text) & "' where id='" & getId(pCombo.Text) & "'"
End Sub

Private Sub Form_Load()
connectdb
'Call bindComboBox(pCombo, "Select * from food")
If rs.State Then rs.Close
rs.Open "select isnull(max(id)+1,1) from food", con
lblId.Caption = rs(0)

If rs.State Then rs.Close
rs.Open "select * from food", con
Set MSHFlexGrid1.DataSource = rs
End Sub

Private Sub MSHFlexGrid1_Click()
lblId.Caption = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0)
nametxt.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
amounttxt.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
stocktxt.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3)
End Sub
