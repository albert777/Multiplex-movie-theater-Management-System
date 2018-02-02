VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form update 
   Caption         =   "Form1"
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   735
      Left            =   3360
      TabIndex        =   13
      Top             =   6000
      Width           =   3735
      Begin VB.ComboBox cmbYy 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Text            =   "Combo1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbDd 
         Height          =   315
         Left            =   2520
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbMm 
         Height          =   315
         Left            =   1320
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label txtyy 
         Caption         =   "Year"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   495
      End
      Begin VB.Label txtdd 
         Caption         =   "Date"
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   120
         Width           =   495
      End
      Begin VB.Label txtmm 
         Caption         =   "Month"
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   120
         Width           =   495
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1935
      Left            =   3360
      TabIndex        =   12
      Top             =   2160
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3413
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
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Height          =   615
      Left            =   3360
      TabIndex        =   10
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CheckBox Check4 
      Caption         =   "4"
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   5280
      Width           =   615
   End
   Begin VB.CheckBox Check3 
      Caption         =   "3"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   5280
      Width           =   615
   End
   Begin VB.CheckBox Check2 
      Caption         =   "2"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   5280
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   3360
      TabIndex        =   5
      Top             =   5040
      Width           =   3735
      Begin VB.CheckBox Check1 
         Caption         =   "1"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.ComboBox screenCombo 
      Height          =   315
      ItemData        =   "update.frx":0000
      Left            =   3360
      List            =   "update.frx":000D
      TabIndex        =   4
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label lblId 
      Caption         =   "InvisibleId"
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
      Left            =   5880
      TabIndex        =   11
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Select Date"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Select Show"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Select Screen"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Select Film"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
   End
End
Attribute VB_Name = "update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim dt As String
dt = cmbMm.Text & "/" & cmbDd.Text & "/" & cmbYy.Text

    If Check1.Value = 1 Then con.Execute ("insert into movieallot values('" & movieid & "','" & screenCombo.Text & "','" & 1 & "','" & dt & "')")
    If Check2.Value = 1 Then con.Execute ("insert into movieallot values('" & movieid & "','" & screenCombo.Text & "','" & 2 & "','" & dt & "')")
    If Check3.Value = 1 Then con.Execute ("insert into movieallot values('" & movieid & "','" & screenCombo.Text & "','" & 3 & "','" & dt & "')")
    If Check4.Value = 1 Then con.Execute ("insert into movieallot values('" & movieid & "','" & screenCombo.Text & "','" & 4 & "','" & dt & "')")
   End Sub

Private Sub Form_Load()
connectdb
If rs.State Then rs.Close
rs.Open "select * from addfilm ", con
Set MSHFlexGrid1.DataSource = rs

End Sub

Private Sub MSHFlexGrid1_Click()
lblId.Caption = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0)
Frame2.Enabled = True
bindDate
End Sub


Sub bindDate()
Dim i  As Integer
For i = 1 To 12
    cmbMm.AddItem (i)
Next i

For i = 1 To 31
    cmbDd.AddItem (i)
Next i

i = Year(DateTime.Date)
While i > 2000
    cmbYy.AddItem (i)
    i = i - 1
Wend

cmbDd.Text = 1
cmbMm.Text = 1
cmbYy.Text = Year(DateTime.Date)
End Sub
