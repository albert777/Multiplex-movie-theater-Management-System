VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form update 
   Caption         =   "Form1"
   ClientHeight    =   8385
   ClientLeft      =   2505
   ClientTop       =   1620
   ClientWidth     =   14370
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   14370
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1935
      Left            =   240
      TabIndex        =   13
      Top             =   2280
      Width           =   6855
      _ExtentX        =   12091
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
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   735
      Left            =   9480
      TabIndex        =   6
      Top             =   3000
      Width           =   3735
      Begin VB.ComboBox cmbYy 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbDd 
         Height          =   315
         Left            =   2520
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbMm 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label txtyy 
         Caption         =   "Year"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   495
      End
      Begin VB.Label txtdd 
         Caption         =   "Date"
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         Top             =   120
         Width           =   495
      End
      Begin VB.Label txtmm 
         Caption         =   "Month"
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000A&
      Caption         =   "Update"
      Height          =   375
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   9480
      TabIndex        =   4
      Top             =   2160
      Width           =   3735
      Begin VB.CheckBox Check4 
         Caption         =   "4"
         Height          =   375
         Left            =   2160
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox Check3 
         Caption         =   "3"
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Caption         =   "2"
         Height          =   375
         Left            =   840
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "1"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblId 
         Caption         =   "0"
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
         Left            =   2880
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.ComboBox screenCombo 
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
      ItemData        =   "ò.frx":0000
      Left            =   9480
      List            =   "ò.frx":000D
      TabIndex        =   3
      Text            =   "1"
      Top             =   3960
      Width           =   2175
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   2535
      Left            =   240
      TabIndex        =   14
      Top             =   5040
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   4471
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
   Begin VB.Label Label4 
      Caption         =   "Select Date"
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Select Show"
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Select Screen"
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   3960
      Width           =   1935
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

'------------------------------------Add Show 1------------------------------------'
If rs.State Then rs.Close
con.Execute "delete from movieallot where movieid='" & lblId.Caption & "' and screenid='" & screenCombo.Text & "' and Show=1"
If Check1.Value = 1 Then con.Execute ("insert into movieallot values('" & lblId.Caption & "','" & screenCombo.Text & "','" & 1 & "','" & dt & "')")

'------------------------------------Add Show 2------------------------------------'
con.Execute "delete from movieallot where movieid='" & lblId.Caption & "' and screenid='" & screenCombo.Text & "' and Show=2"
If Check2.Value = 1 Then con.Execute ("insert into movieallot values('" & lblId.Caption & "','" & screenCombo.Text & "','" & 2 & "','" & dt & "')")

'------------------------------------Add Show 3------------------------------------'
con.Execute "delete from movieallot where movieid='" & lblId.Caption & "' and screenid='" & screenCombo.Text & "' and Show=3"
If Check3.Value = 1 Then con.Execute ("insert into movieallot values('" & lblId.Caption & "','" & screenCombo.Text & "','" & 3 & "','" & dt & "')")

'------------------------------------Add Show 4------------------------------------'
con.Execute "delete from movieallot where movieid='" & lblId.Caption & "' and screenid='" & screenCombo.Text & "' and Show=4"
If Check4.Value = 1 Then con.Execute ("insert into movieallot values('" & lblId.Caption & "','" & screenCombo.Text & "','" & 4 & "','" & dt & "')")
End Sub

Private Sub Form_Load()
connectdb
If rs.State Then rs.Close
rs.Open "select   name as [Choose A Movie From List], language, actor, actress, director, duration, certification,id  from addfilm ", con
Set MSHFlexGrid1.DataSource = rs

End Sub

Private Sub MSHFlexGrid1_Click()
lblId.Caption = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 7)
Frame2.Enabled = True
Call bindDate
End Sub


Sub bindDate()
Dim i  As Integer
For i = 1 To 12
    cmbMm.AddItem (i)
Next i

For i = 1 To 31
    cmbDd.AddItem (i)
Next i

i = Year(DateTime.date)
While i > 2000
    cmbYy.AddItem (i)
    i = i - 1
Wend

cmbDd.Text = 1
cmbMm.Text = 1
cmbYy.Text = Year(DateTime.date)
End Sub

Sub bindGrid()
If rs.State Then rs.Close
rs.Open "SELECT     f.name, f.language, f.actor, f.actress, f.director, f.duration, f.certification, a.screenid, a.show, a.date as ReleaseDate from dbo.addfilm AS f INNER JOIN dbo.movieallot AS a ON f.id = a.movieid where f.id='" & lblId.Caption & "'", con
Set MSHFlexGrid1.DataSource = rs
End Sub
