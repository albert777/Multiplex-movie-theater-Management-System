VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "Msadodc.ocx"
Begin VB.Form login 
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   8475
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3720
      Top             =   3120
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"login.frx":0000
      OLEDBString     =   $"login.frx":008B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton login 
      Caption         =   "Log In"
      Height          =   615
      Left            =   3720
      TabIndex        =   4
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "admin"
      Top             =   1920
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Text            =   "admin"
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
connectdb
End Sub

Private Sub login_Click()
If rs.State Then rs.Close
rs.Open ("select utype from login where username = '" & Text1.Text & "' and password = '" & Text2.Text & "'"), con

If rs.BOF = True Then
    MsgBox ("Invalid")
    Text1.Text = ""
    Text2.Text = ""
    Exit Sub
End If

Dim utype As String
utype = rs(0)
If utype = "staff" Then
    Dim attnCount As Integer
    
    If rs.State Then rs.Close
    rs.Open "select count(*) from attendence where staffUsername ='" & Text1.Text & "' and convert(date,dtTym)=convert(date,getdate())", con
    If Val(rs(0)) <= 0 Then con.Execute "insert into dbo.attendence values('" & Text1.Text & "',getdate())"
    
    If rs.State Then rs.Close
    rs.Open "select count(*) from attendence where staffUsername='" & Text1.Text & "' and convert(date, dtTym)=convert(date,getdate())", con
    attnCount = Val(rs(0))
    If attnCount <= 0 Then con.Execute "insert into attendence values ('" & Text1.Text & "') "
    MDIForm1.smnNewFilm.Visible = False
    MDIForm1.Show
    Unload Me
ElseIf utype = "admin" Then
    MDIForm1.smnNewFilm.Visible = False
    MDIForm1.smnStaff.Visible = False
    MDIForm1.smnSalary.Visible = False
    MDIForm1.smnItemSock.Visible = False
    MDIForm1.Show
    Unload Me
End If
End Sub

