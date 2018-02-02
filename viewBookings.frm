VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form viewBookings 
   Caption         =   "Form1"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15480
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   15480
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   13200
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Include Cancell"
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
      Left            =   11040
      TabIndex        =   9
      Top             =   1320
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4335
      Left            =   480
      TabIndex        =   8
      Top             =   1920
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   7646
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
   Begin VB.ComboBox comboMovie 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3120
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2535
   End
   Begin VB.ComboBox comboscreen 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "viewBookings.frx":0000
      Left            =   8400
      List            =   "viewBookings.frx":0002
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
   Begin VB.ComboBox comboShow 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "viewBookings.frx":0004
      Left            =   5760
      List            =   "viewBookings.frx":0014
      TabIndex        =   0
      Text            =   "1"
      Top             =   1200
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      Format          =   124321795
      CurrentDate     =   42295
   End
   Begin VB.Label date 
      Caption         =   "Date"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Movie Name"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Screen"
      Height          =   255
      Left            =   8400
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Show"
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "viewBookings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub comboMovie_Click()
Call bindComboBox(comboscreen, "select screenid,'' from movieallot where date in(select max(date) from movieallot where date<= '" & DTPicker1.Value & "')")
End Sub

Private Sub Command1_Click()
Dim s As String
If Check1.Value = 1 Then
    s = "SELECT     case(c.bookStatus) when 0 then 'Cancelled' when 1 then 'Confirm' end as Status, f.name, c.seatno, c.amount, m.name AS BookName, m.phone, m.show, m.screen, m.bookingDate from dbo.bookingMaster AS m INNER JOIN dbo.bookingChild AS c ON m.billno = c.billno INNER JOIN dbo.addfilm AS f ON m.filmId = f.id where m.date=convert(date,'" & DTPicker1.Value & "') and m.show='" & comboShow.Text & "' and m.screen='" & getId(comboscreen.Text) & "'"
Else
    s = "SELECT     case(c.bookStatus) when 0 then 'Cancelled' when 1 then 'Confirm' end as Status, f.name, c.seatno, c.amount, m.name AS BookName, m.phone, m.show, m.screen, m.bookingDate from dbo.bookingMaster AS m INNER JOIN dbo.bookingChild AS c ON m.billno = c.billno INNER JOIN dbo.addfilm AS f ON m.filmId = f.id where m.date=convert(date,'" & DTPicker1.Value & "') and m.show='" & comboShow.Text & "' and m.screen='" & getId(comboscreen.Text) & "' and c.bookstatus=1"
End If
If rs.State Then rs.Close
rs.Open s, con
Set MSHFlexGrid1.DataSource = rs
End Sub

Private Sub DTPicker1_Change()
Dim dt As Date
Dim s As String
dt = DTPicker1.Value
s = "select distinct a.id,a.name from movieallot as m inner join addfilm as a on m.movieid=a.id where m.date in(select max(date) from movieallot where date<= '" & DateValue(dt) & "')"
Call bindComboBox(comboMovie, s)
End Sub

Private Sub Form_Load()
connectdb
DTPicker1.Value = Now
'"Using 'mm/dd/yyyy':"; Tab(30); Format$(Now, "mm/dd/yyyy")

End Sub
