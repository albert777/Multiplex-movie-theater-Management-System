VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form tktBooking 
   Caption         =   "Form1"
   ClientHeight    =   8535
   ClientLeft      =   2115
   ClientTop       =   1425
   ClientWidth     =   15975
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8535
   ScaleWidth      =   15975
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3135
      Left            =   0
      TabIndex        =   33
      Top             =   120
      Width           =   5295
      Begin VB.TextBox nametxt 
         Height          =   375
         Left            =   2040
         TabIndex        =   44
         Top             =   2160
         Width           =   3015
      End
      Begin VB.ComboBox comboMovie 
         Height          =   315
         ItemData        =   "tktBooking.frx":0000
         Left            =   2040
         List            =   "tktBooking.frx":0002
         TabIndex        =   37
         Text            =   "Combo1"
         Top             =   720
         Width           =   3135
      End
      Begin VB.ComboBox comboscreen 
         Height          =   315
         ItemData        =   "tktBooking.frx":0004
         Left            =   2040
         List            =   "tktBooking.frx":0006
         TabIndex        =   36
         Top             =   1680
         Width           =   3135
      End
      Begin VB.ComboBox comboShow 
         Height          =   315
         ItemData        =   "tktBooking.frx":0008
         Left            =   2040
         List            =   "tktBooking.frx":000A
         TabIndex        =   35
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox phonetxt 
         Height          =   375
         Left            =   2040
         TabIndex        =   34
         Top             =   2640
         Width           =   3015
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
         Left            =   2040
         TabIndex        =   38
         Top             =   120
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   83492867
         CurrentDate     =   42295
      End
      Begin VB.Label Label7 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label date 
         Caption         =   "Date"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Movie Name"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Screen"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Show"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Phone"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   2640
         Width           =   1455
      End
   End
   Begin VB.TextBox txtRemoveMsg 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000D&
      Height          =   165
      Left            =   2640
      TabIndex        =   32
      Text            =   "Click on SeatNo. to remove"
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controls Used for coding"
      Height          =   735
      Left            =   0
      TabIndex        =   27
      Top             =   4920
      Visible         =   0   'False
      Width           =   1935
      Begin VB.TextBox seattxt 
         Height          =   375
         Left            =   600
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ALLOT"
         Height          =   375
         Left            =   720
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   " seat number"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000E&
      Caption         =   "BOOK"
      Height          =   495
      Left            =   2040
      MaskColor       =   &H00808080&
      TabIndex        =   3
      Top             =   7440
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   8055
      Left            =   5520
      Picture         =   "tktBooking.frx":000C
      ScaleHeight     =   7995
      ScaleWidth      =   9915
      TabIndex        =   2
      Top             =   0
      Width           =   9975
      Begin VB.Label lbls22 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "22"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   9360
         TabIndex        =   26
         Top             =   7080
         Width           =   495
      End
      Begin VB.Label lbls21 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   8160
         TabIndex        =   25
         Top             =   7080
         Width           =   495
      End
      Begin VB.Label lbls20 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   6960
         TabIndex        =   24
         Top             =   7080
         Width           =   495
      End
      Begin VB.Label lbls19 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   5880
         TabIndex        =   23
         Top             =   7080
         Width           =   495
      End
      Begin VB.Label lbls18 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   3480
         TabIndex        =   22
         Top             =   7080
         Width           =   495
      End
      Begin VB.Label lbls17 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   2400
         TabIndex        =   21
         Top             =   7080
         Width           =   495
      End
      Begin VB.Label lbls16 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   1200
         TabIndex        =   20
         Top             =   7080
         Width           =   495
      End
      Begin VB.Label lbls15 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   120
         TabIndex        =   19
         Top             =   7080
         Width           =   495
      End
      Begin VB.Label lbls14 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   8880
         TabIndex        =   18
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label lbls13 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   7920
         TabIndex        =   17
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label lbls12 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   6840
         TabIndex        =   16
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label lbls11 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   5640
         TabIndex        =   15
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label lbls10 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   3720
         TabIndex        =   14
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label lbls9 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   2640
         TabIndex        =   13
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label lbls8 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   1560
         TabIndex        =   12
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label lbls7 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   600
         TabIndex        =   11
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label lbls6 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   7560
         TabIndex        =   10
         Top             =   5160
         Width           =   495
      End
      Begin VB.Label lbls5 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   6600
         TabIndex        =   9
         Top             =   5160
         Width           =   495
      End
      Begin VB.Label lbls4 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   5520
         TabIndex        =   8
         Top             =   5160
         Width           =   495
      End
      Begin VB.Label lbls3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   3960
         TabIndex        =   7
         Top             =   5160
         Width           =   495
      End
      Begin VB.Label lbls2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   2880
         TabIndex        =   6
         Top             =   5160
         Width           =   495
      End
      Begin VB.Label lbls1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   1800
         TabIndex        =   5
         Top             =   5160
         Width           =   495
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2895
      Left            =   2040
      TabIndex        =   31
      Top             =   3360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   5106
      _Version        =   393216
      Rows            =   5
      Cols            =   1
      FixedCols       =   0
      RowHeightMin    =   500
      FormatString    =   "Selected Seat Numbers"
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
      _Band(0).Cols   =   1
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Total Amount"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblSeatAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   2040
      TabIndex        =   0
      Top             =   6240
      Width           =   3135
   End
End
Attribute VB_Name = "tktBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub comboMovie_Click()
Call bindComboBox(comboShow, "select show,'' from movieallot where movieid='" & getId(comboMovie.Text) & "' and date in(select max(date) from movieallot where date<= '" & DTPicker1.Value & "')")
End Sub

Private Sub comboscreen_Click()
Dim bookDt As Date
Dim i As Integer
Dim movieid As Integer
Dim lblSeat As Label
Dim qry As String

movieid = getId(comboMovie.Text)
bookDt = DTPicker1.Value

If rs.State Then rs.Close
rs.Open "select seatno from bookingChild as c inner join bookingMaster as m on c.billno=m.billno where m.filmid = '" & movieid & "' and m.Show = '" & getId(comboShow.Text) & "'and m.Screen= '" & getId(comboscreen.Text) & "'and m.[date]= '" & DTPicker1.Value & "' and c.bookStatus = 1"

If rs.BOF = False Then
    rs.MoveFirst
    While Not rs.EOF
        If rs.Fields(0) = 1 Then lbls1.Enabled = False
        If rs.Fields(0) = 2 Then lbls2.Enabled = False
        If rs.Fields(0) = 3 Then lbls3.Enabled = False
        If rs.Fields(0) = 4 Then lbls4.Enabled = False
        If rs.Fields(0) = 5 Then lbls5.Enabled = False
        If rs.Fields(0) = 6 Then lbls6.Enabled = False
        If rs.Fields(0) = 7 Then lbls7.Enabled = False
        If rs.Fields(0) = 8 Then lbls8.Enabled = False
        If rs.Fields(0) = 9 Then lbls9.Enabled = False
        If rs.Fields(0) = 10 Then lbls10.Enabled = False
        If rs.Fields(0) = 11 Then lbls11.Enabled = False
        If rs.Fields(0) = 12 Then lbls12.Enabled = False
        If rs.Fields(0) = 13 Then lbls13.Enabled = False
        If rs.Fields(0) = 14 Then lbls14.Enabled = False
        If rs.Fields(0) = 15 Then lbls15.Enabled = False
        If rs.Fields(0) = 16 Then lbls16.Enabled = False
        If rs.Fields(0) = 17 Then lbls17.Enabled = False
        If rs.Fields(0) = 18 Then lbls18.Enabled = False
        If rs.Fields(0) = 19 Then lbls19.Enabled = False
        If rs.Fields(0) = 20 Then lbls20.Enabled = False
        If rs.Fields(0) = 21 Then lbls21.Enabled = False
        If rs.Fields(0) = 22 Then lbls22.Enabled = False
        rs.MoveNext
    Wend
End If

'For i = 1 To 22
'
'qry = "select count(*) from bookingChild as c inner join bookingMaster as m on c.billno=m.billno where c.filmid = '" & movieid & "' and c.Show = '" & comboShow.Text & "'and c.Screen= '" & comboscreen.Text & "'and c.SeatNo= '" & i & "'and m.[date]= '" & DTPicker1.Value & "' and c.bookStatus = 1"
'
'
'
'    If rs.State Then rs.Close
'    rs.Open qry, con
'    If Val(rs(0)) > 0 Then
'        If i = 1 Then lbls1.Enabled = False
'        If i = 2 Then lbls2.Enabled = False
'        If i = 3 Then lbls3.Enabled = False
'        If i = 4 Then lbls4.Enabled = False
'        If i = 5 Then lbls5.Enabled = False
'        If i = 6 Then lbls6.Enabled = False
'        If i = 7 Then lbls7.Enabled = False
'        If i = 8 Then lbls8.Enabled = False
'        If i = 9 Then lbls9.Enabled = False
'        If i = 10 Then lbls10.Enabled = False
'        If i = 11 Then lbls11.Enabled = False
'        If i = 12 Then lbls12.Enabled = False
'        If i = 13 Then lbls13.Enabled = False
'        If i = 14 Then lbls14.Enabled = False
'        If i = 15 Then lbls15.Enabled = False
'        If i = 16 Then lbls16.Enabled = False
'        If i = 17 Then lbls17.Enabled = False
'        If i = 18 Then lbls18.Enabled = False
'        If i = 19 Then lbls19.Enabled = False
'        If i = 20 Then lbls20.Enabled = False
'        If i = 21 Then lbls21.Enabled = False
'        If i = 22 Then lbls22.Enabled = False
'    End If
'Next i
End Sub


Private Sub comboShow_Click()
Call bindComboBox(comboscreen, "select screenid,'' from movieallot where movieid='" & getId(comboMovie.Text) & "' and show='" & getId(comboShow.Text) & "' and date in(select max(date) from movieallot where date<= '" & DTPicker1.Value & "')")
End Sub

Private Sub Command1_Click()
'Label7.Caption = Val(Label7.Caption) + 1
'
'Dim str As String
'listtxt.AddItem (seattxt.Text)
'seattxt = ""
'If rs.State Then rs.Close
'rs.Open "select amount from tktrate where screen = '" & getId(comboscreen.Text) & "' and show = '" & comboShow.Text & "'", con
'lblSeatAmt.Caption = Val(lblSeatAmt.Caption) + Val(rs(0))



End Sub

Private Sub Command2_Click()
Dim maxBill As Integer
Dim i As Integer
Dim str As String
Dim amtPerSeat As Double

If rs.State Then rs.Close
rs.Open "select isnull(amount,0) from tktrate where screen = '" & getId(comboscreen.Text) & "' and show = '" & getId(comboShow.Text) & "'", con
amtPerSeat = rs(0)
If amtPerSeat = 0 Then
    MsgBox "Please add Price for selected Show on Selected Screen"
    Exit Sub
End If

con.Execute " insert into bookingmaster values('" & DTPicker1.Value & "','" & nametxt.Text & "','" & phonetxt.Text & "','" & getId(comboMovie.Text) & "','" & getId(comboShow.Text) & "','" & getId(comboscreen.Text) & "',getdate())"

If rs.State Then rs.Close
rs.Open "select isnull(max(billno),0) from bookingmaster", con
maxBill = rs(0)

For i = 1 To 4
    If MSHFlexGrid1.TextMatrix(i, 0) <> "" Then
        str = "insert into bookingchild values('" & maxBill & "','" & MSHFlexGrid1.TextMatrix(i, 0) & "','" & amtPerSeat & "','1')"
        con.Execute str
    End If
Next i
End Sub

Private Sub DTPicker1_Change()
Call bindComboBox(comboMovie, "select distinct a.id,a.name from movieallot as m inner join addfilm as a on m.movieid=a.id where m.date in(select max(date) from movieallot where date<= '" & DTPicker1.Value & "')")
comboscreen.clear
End Sub

Private Sub Form_Load()
connectdb
'"Using 'mm/dd/yyyy':"; Tab(30); Format$(Now, "mm/dd/yyyy")
Dim dt As Date
Dim s As String
MSHFlexGrid1.RowHeight(0) = 800
txtRemoveMsg.Top = MSHFlexGrid1.Top + MSHFlexGrid1.CellTop - MSHFlexGrid1.CellHeight + txtRemoveMsg.Height + 50
DTPicker1.Value = Now
dt = DTPicker1.Value
s = "select m.movieid,a.name from movieallot as m inner join addfilm as a on m.movieid=a.id where m.date in(select max(date) from movieallot where date<= '" & dt & "')"
Call bindComboBox(comboMovie, s)
End Sub

Private Function bindBookedSeats(ByVal seatNo As Integer)
Dim i As Integer
Dim flag As Integer
Dim billAmt As Double
Dim tktRate As Double

flag = 0
If rs.State Then rs.Close
rs.Open "select isnull(amount,0) from tktrate where screen = '" & getId(comboscreen.Text) & "' and show = '" & getId(comboShow.Text) & "'", con
If rs.BOF = False Then tktRate = rs(0)
If tktRate <= 0 Then
    MsgBox "Please add Price for selected Show on Selected Screen"
    Exit Function
End If

For i = 1 To MSHFlexGrid1.Rows - 1
    If Val(MSHFlexGrid1.TextMatrix(i, 0)) = seatNo Or flag = 1 Then Exit Function
    If MSHFlexGrid1.TextMatrix(i, 0) = "" And flag = 0 Then
        MSHFlexGrid1.TextMatrix(i, 0) = seatNo
        '----------Add Amount----------
        billAmt = Val(lblSeatAmt.Caption)
        lblSeatAmt.Caption = billAmt + tktRate
        '----------Add Amount----------
        flag = 1
    End If
Next i
If flag = 0 Then MsgBox "You have reached maximum Seats"
End Function


'--------------------Seat Selections Starts--------------------'
Private Sub lbls1_Click()
Call bindBookedSeats(lbls1.Caption)
End Sub

Private Sub lbls10_Click()
Call bindBookedSeats(lbls10.Caption)

End Sub

Private Sub lbls11_Click()
Call bindBookedSeats(lbls11.Caption)

End Sub

Private Sub lbls12_Click()
Call bindBookedSeats(lbls12.Caption)

End Sub

Private Sub lbls13_Click()
Call bindBookedSeats(lbls13.Caption)


End Sub

Private Sub lbls14_Click()
Call bindBookedSeats(lbls14.Caption)


End Sub

Private Sub lbls15_Click()
Call bindBookedSeats(lbls15.Caption)


End Sub

Private Sub lbls16_Click()
Call bindBookedSeats(lbls16.Caption)


End Sub

Private Sub lbls17_Click()
Call bindBookedSeats(lbls17.Caption)

End Sub

Private Sub lbls18_Click()
Call bindBookedSeats(lbls18.Caption)


End Sub

Private Sub lbls19_Click()
Call bindBookedSeats(lbls19.Caption)

End Sub

Private Sub lbls2_Click()
Call bindBookedSeats(lbls2.Caption)

End Sub

Private Sub lbls20_Click()
Call bindBookedSeats(lbls20.Caption)


End Sub

Private Sub lbls21_Click()
Call bindBookedSeats(lbls21.Caption)


End Sub

Private Sub lbls22_Click()
Call bindBookedSeats(lbls22.Caption)


End Sub

Private Sub lbls3_Click()
Call bindBookedSeats(lbls3.Caption)


End Sub

Private Sub lbls4_Click()
Call bindBookedSeats(lbls4.Caption)

End Sub

Private Sub lbls5_Click()
Call bindBookedSeats(lbls5.Caption)


End Sub

Private Sub lbls6_Click()
Call bindBookedSeats(lbls6.Caption)

End Sub

Private Sub lbls7_Click()
Call bindBookedSeats(lbls7.Caption)

End Sub

Private Sub lbls8_Click()
Call bindBookedSeats(lbls8.Caption)

End Sub

Private Sub lbls9_Click()
Call bindBookedSeats(lbls9.Caption)
            
End Sub
'--------------------Seat Selections Ends--------------------'
Private Sub MSHFlexGrid1_Click()
If MsgBox("Do you want to remove selected seat?", vbYesNo, "Replace") = vbYes Then
    MSHFlexGrid1.RemoveItem (MSHFlexGrid1.Row)
    Call MSHFlexGrid1.AddItem("", MSHFlexGrid1.Rows)
    '----------Deduct Amount----------
    If rs.State Then rs.Close
    rs.Open "select isnull(amount,0) from tktrate where screen = '" & getId(comboscreen.Text) & "' and show = '" & getId(comboShow.Text) & "'", con
    
    billAmt = Val(lblSeatAmt.Caption)
    If rs.BOF = False Then tktRate = rs(0)
    
    lblSeatAmt.Caption = billAmt - tktRate
    '----------Deduct Amount----------
End If
End Sub

