VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form rate 
   Caption         =   "s"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   14910
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2295
      Left            =   7560
      TabIndex        =   11
      Top             =   2040
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4048
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
      Caption         =   "ADD"
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox amounttxt 
      Height          =   495
      Left            =   3720
      TabIndex        =   9
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CheckBox Check4 
      Caption         =   "4"
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   2760
      Width           =   615
   End
   Begin VB.CheckBox Check3 
      Caption         =   "3"
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   2760
      Width           =   615
   End
   Begin VB.CheckBox Check2 
      Caption         =   "2"
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   2760
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   3720
      TabIndex        =   4
      Top             =   2640
      Width           =   3135
      Begin VB.CheckBox Check1 
         Caption         =   "1"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.ComboBox screenCombo 
      Height          =   315
      ItemData        =   "Multiplex.frx":0000
      Left            =   3720
      List            =   "Multiplex.frx":000D
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Amount"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Select Show"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Select Screen"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   2040
      Width           =   2055
   End
End
Attribute VB_Name = "rate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'------------------------------------Add Rate for Show 1------------------------------------'
If rs.State Then rs.Close
rs.Open "select count(*) from tktrate where screen='" & screenCombo.Text & "' and show=1", con
If Val(rs(0)) <= 0 Then
    If Check1.Value = 1 Then con.Execute ("insert into tktrate values('" & screenCombo.Text & "','" & 1 & "','" & amounttxt.Text & "')")
Else
    If Check1.Value = 1 Then con.Execute "update amount where screen='" & screenCombo.Text & "' and show=1"
End If

'------------------------------------Add Rate for Show 2------------------------------------'
If rs.State Then rs.Close
rs.Open "select count(*) from tktrate where screen='" & screenCombo.Text & "' and show=2", con
If Val(rs(0)) <= 0 Then
    If Check2.Value = 1 Then con.Execute ("insert into tktrate values('" & screenCombo.Text & "','" & 2 & "','" & amounttxt.Text & "')")
Else
    If Check2.Value = 1 Then con.Execute "update tktrate set amount where screen='" & screenCombo.Text & "' and show=2"
End If

'------------------------------------Add Rate for Show 3------------------------------------'
If rs.State Then rs.Close
rs.Open "select count(*) from tktrate where screen='" & screenCombo.Text & "' and show=3", con
If Val(rs(0)) <= 0 Then
    If Check3.Value = 1 Then con.Execute ("insert into tktrate values('" & screenCombo.Text & "','" & 3 & "','" & amounttxt.Text & "')")
Else
    If Check3.Value = 1 Then con.Execute "update amount where screen='" & screenCombo.Text & "' and show=3"
End If

'------------------------------------Add Rate for Show 4------------------------------------'
If rs.State Then rs.Close
rs.Open "select count(*) from tktrate where screen='" & screenCombo.Text & "' and show=4", con
If Val(rs(0)) <= 0 Then
    If Check4.Value = 1 Then con.Execute ("insert into tktrate values('" & screenCombo.Text & "','" & 4 & "','" & amounttxt.Text & "')")
Else
    If Check4.Value = 1 Then con.Execute "update amount where screen='" & screenCombo.Text & "' and show=4"
End If

screenCombo.clear
amounttxt.Text = ""



End Sub

Private Sub Form_Load()
connectdb
End Sub
Sub bindGrid()
If rs.State Then rs.Close
rs.Open "select * from dbo.tktrate", con
Set MSHFlexGrid1.DataSource = rs
End Sub

