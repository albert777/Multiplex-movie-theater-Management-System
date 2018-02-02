VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form addfilm 
   Caption         =   "Form1"
   ClientHeight    =   8400
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   16575
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   16575
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check5 
      Caption         =   "Include Old Films"
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
      Left            =   7680
      TabIndex        =   27
      Top             =   1440
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   615
      Left            =   5280
      TabIndex        =   25
      Top             =   7680
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3600
      TabIndex        =   24
      Top             =   6720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   119537665
      CurrentDate     =   42322
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
      ItemData        =   "Form2.frx":0000
      Left            =   3600
      List            =   "Form2.frx":000D
      TabIndex        =   23
      Text            =   "1"
      Top             =   5520
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   3600
      TabIndex        =   18
      Top             =   6000
      Width           =   3855
      Begin VB.CheckBox Check4 
         Caption         =   "4"
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox Check3 
         Caption         =   "3"
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox Check2 
         Caption         =   "2"
         Height          =   255
         Left            =   1080
         TabIndex        =   20
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   615
      Left            =   3600
      TabIndex        =   17
      Top             =   7680
      Width           =   1575
   End
   Begin VB.ComboBox certificationCombo 
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
      Left            =   3600
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   5040
      Width           =   3855
   End
   Begin VB.ComboBox directorCombo 
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
      Left            =   3600
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   3840
      Width           =   3855
   End
   Begin VB.ComboBox actresscombo 
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
      Left            =   3600
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   3240
      Width           =   3855
   End
   Begin VB.ComboBox actorCombo 
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
      Left            =   3600
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   2640
      Width           =   3855
   End
   Begin VB.ComboBox languageCombo 
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
      Left            =   3600
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   2040
      Width           =   3855
   End
   Begin VB.TextBox Txtduration 
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   4320
      Width           =   3855
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   1320
      Width           =   3855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5055
      Left            =   7680
      TabIndex        =   26
      Top             =   2040
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   8916
      _Version        =   393216
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblId 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   28
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "From Date"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   6840
      Width           =   2535
   End
   Begin VB.Label Label9 
      Caption         =   "Mark Shows"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   6240
      Width           =   2535
   End
   Begin VB.Label Label8 
      Caption         =   "Choose Screen"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label Label7 
      Caption         =   "Certification"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Enter Duration"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Enter or Choose Director"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Enter or Choose Actress"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Enter or Choose Actor"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Language"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Film Name"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   2535
   End
End
Attribute VB_Name = "addfilm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check5_Click()
If Check5.Value = 1 Then
    If rs.State Then rs.Close
    rs.Open "select * from addfilm", con
    Set MSHFlexGrid1.DataSource = rs
Else
    Call bindGrid
End If
End Sub

Private Sub Command1_Click()
If chkDataExistency("addfilm", "id", lblId.Caption) > 0 Then Exit Sub 'If ID exists

If rs.State Then rs.Close
rs.Open "select count(*) from addfilm where name='" & txtName.Text & "' and language='" & languageCombo.Text & "'", con, adOpenDynamic, adLockOptimistic
If Val(rs(0)) > 0 Then
    MsgBox "Movie Already added"
    Exit Sub
End If

Dim str As String
str = "insert into addfilm values('" & txtName.Text & "','" & languageCombo.Text & "','" & actorCombo.Text & "','" & actresscombo.Text & "','" & directorCombo.Text & "','" & Txtduration.Text & "','" & certificationCombo.Text & "')"
'Txtduration.Text = str
con.Execute str
If rs.State Then rs.Close
rs.Open " select max(id) from addfilm", con
If rs.BOF = False Then
Dim movieid As Integer
    movieid = rs(0)
    '-----------Allot Movie-----------'
    Dim s1, s3 As String
    If Check1.Value = 1 Then con.Execute ("insert into movieallot values('" & movieid & "','" & screenCombo.Text & "','" & 1 & "','" & DTPicker1.Value & "')")
    If Check2.Value = 1 Then con.Execute ("insert into movieallot values('" & movieid & "','" & screenCombo.Text & "','" & 2 & "','" & DTPicker1.Value & "')")
    If Check3.Value = 1 Then con.Execute ("insert into movieallot values('" & movieid & "','" & screenCombo.Text & "','" & 3 & "','" & DTPicker1.Value & "')")
    If Check4.Value = 1 Then con.Execute ("insert into movieallot values('" & movieid & "','" & screenCombo.Text & "','" & 4 & "','" & DTPicker1.Value & "')")
Else
    MsgBox ("Failed in adding movie")
    Exit Sub
End If
Call bindFilmPropery
Call Form_Load
End Sub

Sub bindFilmPropery()
languageCombo.clear
actorCombo.clear
actresscombo.clear
directorCombo.clear

If rs.State Then rs.Close
rs.Open "select distinct language from addfilm order by language  asc", con, adOpenDynamic, adLockOptimistic
If rs.BOF = False Then
    rs.MoveFirst
    languageCombo.Text = rs(0)
    While Not rs.EOF
        languageCombo.AddItem (rs(0))
        rs.MoveNext
    Wend
End If

If rs.State Then rs.Close
rs.Open "select distinct actor from addfilm order by actor asc", con, adOpenDynamic, adLockOptimistic
If rs.BOF = False Then
    rs.MoveFirst
    actorCombo.Text = rs(0)
    While Not rs.EOF
        actorCombo.AddItem (rs(0))
        rs.MoveNext
    Wend
End If

If rs.State Then rs.Close
rs.Open "select distinct actress from addfilm order by actress asc", con, adOpenDynamic, adLockOptimistic
If rs.BOF = False Then
    rs.MoveFirst
    actresscombo.Text = rs(0)
    While Not rs.EOF
        actresscombo.AddItem (rs(0))
        rs.MoveNext
    Wend
End If


If rs.State Then rs.Close
rs.Open "select distinct director from addfilm order by director asc", con, adOpenDynamic, adLockOptimistic
If rs.BOF = False Then
    rs.MoveFirst
    directorCombo.Text = rs(0)
    While Not rs.EOF
        directorCombo.AddItem (rs(0))
        rs.MoveNext
    Wend
End If
End Sub
Sub bindGrid()
If rs.State Then rs.Close
rs.Open "select distinct a.id,a.name,a.language ,a.actor,a.actress,a.director,a.duration,a.certification from movieallot as m inner join addfilm as a on m.movieid=a.id where m.date in(select max(date) from movieallot where date<= '" & Now & "')", con
Set MSHFlexGrid1.DataSource = rs
End Sub

Sub maxId()
If rs.State Then rs.Close
rs.Open "select isnull(max(id)+1,1) from addfilm", con
lblId.Caption = rs(0)
End Sub


Private Sub Command2_Click()
con.Execute "delete from dbo.movieallot where screenid='" & screenCombo.Text & "' and movieid='" & lblId.Caption & "'"
If Check1.Value = 1 Then con.Execute ("insert into movieallot values('" & lblId.Caption & "','" & screenCombo.Text & "','" & 1 & "','" & DTPicker1.Value & "')")
If Check2.Value = 1 Then con.Execute ("insert into movieallot values('" & lblId.Caption & "','" & screenCombo.Text & "','" & 2 & "','" & DTPicker1.Value & "')")
If Check3.Value = 1 Then con.Execute ("insert into movieallot values('" & lblId.Caption & "','" & screenCombo.Text & "','" & 3 & "','" & DTPicker1.Value & "')")
If Check4.Value = 1 Then con.Execute ("insert into movieallot values('" & lblId.Caption & "','" & screenCombo.Text & "','" & 4 & "','" & DTPicker1.Value & "')")
con.Execute "update addfilm set name = '" & txtName.Text & "', language = '" & languageCombo.Text & "', actor = '" & actorCombo.Text & "', actress = '" & actresscombo.Text & "', director = '" & directorCombo.Text & "', duration = '" & Txtduration.Text & "', certification = '" & certificationCombo.Text & "' where id = '" & lblId.Caption & "'"
If Check5.Value = 1 Then Call Check5_Click Else Call Form_Load
End Sub

Private Sub DTPicker1_Change()
If DateValue(DTPicker1.Value) < Now Then DTPicker1.Value = DateAdd("d", 1, Now)
End Sub

Private Sub Form_Load()
connectdb
DTPicker1.Value = DateAdd("d", 1, Now)
Call bindFilmPropery
Call bindGrid
Call maxId
End Sub

Private Sub languageCombo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then actorCombo.SetFocus
End Sub

Private Sub MSHFlexGrid1_Click()
lblId.Caption = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0)
txtName.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
languageCombo.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
actorCombo.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3)
actresscombo.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4)
directorCombo.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 5)
Txtduration.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 6)
certificationCombo.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 7)

Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0

'Call bindComboBox(screenCombo, "select  distinct screenid,'' from dbo.movieallot where movieid='" & lblId.Caption & "'")

If rs.State Then rs.Close
rs.Open "select show from dbo.movieallot where movieid='" & lblId.Caption & "' and screenid='" & screenCombo.Text & "'", con
If rs.BOF = False Then
    rs.MoveFirst
    While Not rs.EOF
        If Val(rs(0)) = 1 Then Check1.Value = 1
        If Val(rs(0)) = 2 Then Check2.Value = 1
        If Val(rs(0)) = 3 Then Check3.Value = 1
        If Val(rs(0)) = 4 Then Check4.Value = 1
        rs.MoveNext
    Wend
End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then languageCombo.SetFocus
End Sub

Private Sub actorCombo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then actresscombo.SetFocus
End Sub

Private Sub actresscombo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then directorCombo.SetFocus
End Sub

Private Sub directorCombo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txtduration.SetFocus
End Sub

