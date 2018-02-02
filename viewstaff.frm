VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form viewstaff 
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15690
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   15690
   Begin VB.CommandButton command1 
      Caption         =   "SEARCH"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox searchtxt 
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3015
      Left            =   720
      TabIndex        =   0
      Top             =   1680
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   5318
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "Click anyone to update"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1320
      Width           =   9375
   End
End
Attribute VB_Name = "viewstaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If rs.State Then rs.Close
rs.Open "select staffid,name,address,dob,qualification,phone,email,salary from staffreg where name = '" & searchtxt.Text & "'", con
Set MSHFlexGrid1.DataSource = rs
End Sub

Private Sub Form_Load()
connectdb
If rs.State Then rs.Close
rs.Open "select staffid,name,address,dob,qualification,phone,email,salary from staffreg", con
Set MSHFlexGrid1.DataSource = rs
End Sub

Private Sub MSHFlexGrid1_Click()
staffreg.lblId.Caption = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0)
staffreg.nametxt.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
staffreg.addresstxt.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
staffreg.dobtxt.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3)
staffreg.qualificationtxt.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4)
staffreg.phonetxt.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 5)
staffreg.emailtxt.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 6)
staffreg.salarytxt.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 7)
staffreg.usernametxt.Visible = False
staffreg.passwordtxt.Visible = False
staffreg.confirmtxt.Visible = False
staffreg.Show
Unload Me
End Sub
