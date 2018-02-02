VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form cancellation 
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14160
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   14160
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   495
      Left            =   5400
      TabIndex        =   8
      Top             =   5280
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2415
      Left            =   1200
      TabIndex        =   7
      Top             =   2760
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4260
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
   Begin VB.ComboBox moviecombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1800
      Width           =   3615
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1320
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      Format          =   83492865
      CurrentDate     =   42281
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox phonetxt 
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label lblAmtToPay 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   5400
      Width           =   1215
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
      Left            =   7080
      TabIndex        =   9
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Phone Number"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Choose Moive"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   1800
      Width           =   1935
   End
End
Attribute VB_Name = "cancellation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
            
Private Sub Command1_Click()
Dim phone As String
Dim s As String
Dim amt As Double

s = "SELECT    c.seatno , CASE (c.bookStatus) WHEN 1 THEN 'Active' WHEN 0 THEN 'Cancelled' END AS statText, m.show, m.screen, f.name, f.language, m.billno, m.date, c.amount AS AmountToPay, c.bookStatus, c.id from dbo.bookingMaster AS m INNER JOIN dbo.bookingChild AS c ON m.billno = c.billno INNER JOIN dbo.addfilm AS f ON m.filmId = f.id where m.phone='" & phonetxt.Text & "' or m.name='" & phonetxt.Text & "' and m.date='" & DateValue(DTPicker1.Value) & "' and m.filmId ='" & getId(moviecombo.Text) & "'"
If rs.State Then rs.Close
rs.Open s, con
Set MSHFlexGrid1.DataSource = rs

amt = 0
For i = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(i, 9) = 1 Then amt = amt + Val(MSHFlexGrid1.TextMatrix(i, 8))
Next i
lblAmtToPay.Caption = amt & ".00"
End Sub

Private Sub Command2_Click()
Dim cid As String
Dim s As String
If lblId.Caption = 0 Then
    MsgBox "Please choose a seat"
    Exit Sub
End If

cid = lblId.Caption
con.Execute "update bookingchild set bookstatus= 0 where id='" & cid & "'"
Call Command1_Click
Set cancellation = Nothing
cancellation.Show
Unload Me
End Sub

Private Sub DTPicker1_Click()
Call bindComboBox(moviecombo, "select m.movieid,a.name from movieallot as m inner join addfilm as a on m.movieid=a.id where m.date in(select max(date) from movieallot where date<= '" & DTPicker1.Value & "')")
End Sub

Private Sub Form_Load()
connectdb
End Sub

Private Sub Label4_Click()

End Sub

Private Sub moviecombo_click()

'If rs.State Then rs.Close
'rs.Open "select f.name, f.language, m.billno, m.date, c.amount, m.phone from dbo.bookingMaster AS m INNER JOIN dbo.bookingChild AS c ON m.billno = c.billno INNER JOIN dbo.addfilm AS f ON m.filmid = f.id where m.filmid='" & getId(moviecombo.Text) & "'", con
'Set MSHFlexGrid1.DataSource = rs
End Sub


Private Sub MSHFlexGrid1_Click()
If MSHFlexGrid1.Rows <= 1 Then Exit Sub
MSHFlexGrid1.clear
Call Command1_Click
For i = 0 To MSHFlexGrid1.Cols - 1
    MSHFlexGrid1.Col = i
    MSHFlexGrid1.CellBackColor = vbHighlight
Next i
MSHFlexGrid1.Col = 0
MSHFlexGrid1.CellBackColor = vbBlue
lblId.Caption = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 10)
End Sub
