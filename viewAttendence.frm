VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form viewAttendence 
   Caption         =   "Form1"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13830
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   13830
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   110886914
      CurrentDate     =   42331
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   2655
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3135
      Left            =   2880
      TabIndex        =   0
      Top             =   2520
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   5530
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
Attribute VB_Name = "viewAttendence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim attnCount As Integer
Dim sal As Double
Dim tot As Double
Call MonthView1_DateClick(DateValue(MonthView1.Value))
MSHFlexGrid1.Cols = MSHFlexGrid1.Cols + 1
For i = 1 To MSHFlexGrid1.Rows - 1
    If rs.State Then rs.Close
    rs.Open "select count(*) from attendence where staffUsername='" & MSHFlexGrid1.TextMatrix(i, 1) & "' and month(dttym) =month('" & MonthView1.Value & "')", con
    attnCount = Val(rs(0))
    If rs.State Then rs.Close
    rs.Open "select salary from dbo.staffreg where username='" & MSHFlexGrid1.TextMatrix(i, 1) & "'", con
    sal = Val(rs(0))
    tot = attnCount * sal
    MSHFlexGrid1.TextMatrix(i, 8) = tot
Next i
End Sub


Private Sub Form_Load()
MonthView1.Value = Now
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
If rs.State Then rs.Close
rs.Open "select s.name,  s.username, s.address, s.dob, s.qualification, s.phone, s.email, a.dtTym from dbo.attendence AS a INNER JOIN dbo.staffreg AS s ON a.staffUsername = s.username where convert(date,a.dttym)=convert(date,'" & MonthView1.Value & "')", con
Set MSHFlexGrid1.DataSource = rs
End Sub
