VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "Msadodc.ocx"
Begin VB.Form staffreg 
   Caption         =   "Form1"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   14970
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   615
      Left            =   5160
      TabIndex        =   22
      Top             =   7680
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   7800
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
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
      Connect         =   $"Form1.frx":0000
      OLEDBString     =   $"Form1.frx":008B
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
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   615
      Left            =   2400
      TabIndex        =   20
      Top             =   7680
      Width           =   2415
   End
   Begin VB.TextBox confirmtxt 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   19
      Top             =   6600
      Width           =   2655
   End
   Begin VB.TextBox passwordtxt 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   17
      Top             =   6000
      Width           =   2655
   End
   Begin VB.TextBox usernametxt 
      Height          =   495
      Left            =   3720
      TabIndex        =   15
      Top             =   5400
      Width           =   2655
   End
   Begin VB.TextBox salarytxt 
      Height          =   495
      Left            =   3720
      TabIndex        =   13
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox emailtxt 
      Height          =   495
      Left            =   3720
      TabIndex        =   11
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox phonetxt 
      Height          =   495
      Left            =   3720
      TabIndex        =   9
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox qualificationtxt 
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox dobtxt 
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox addresstxt 
      Height          =   1095
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox nametxt 
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   600
      Width           =   2655
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
      Height          =   615
      Left            =   6600
      TabIndex        =   21
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Confirm Password"
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   6720
      Width           =   2535
   End
   Begin VB.Label Label9 
      Caption         =   "Password"
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label Label8 
      Caption         =   "Username"
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   5520
      Width           =   2535
   End
   Begin VB.Label Label7 
      Caption         =   "Salary / Day"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "E-mail"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Phone No"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Qualification"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Date of Birth"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Address"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Staff Name"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
End
Attribute VB_Name = "staffreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

con.Execute ("insert into login values ('" & usernametxt.Text & "','" & passwordtxt.Text & "','staff')")
con.Execute ("insert into staffreg values ('" & nametxt.Text & "','" & addresstxt.Text & "','" & dobtxt.Text & "','" & qualificationtxt.Text & "','" & phonetxt.Text & "','" & emailtxt.Text & "','" & salarytxt.Text & "''" & usernametxt.Text & "')")
Call clear
End Sub



Private Sub Command2_Click()
con.Execute "update staffreg set name = '" & txtName.Text & "', address = '" & addresstxt.Text & "', dob = '" & dobtxt.Text & "', qualification = '" & qualificationtxt.Text & "', phone = '" & phonetxt.Text & "', email = '" & emailtxt.Text & "', salary = '" & salarytxt.Text & "',username = '" & usernametxt.Text & "' where id = '" & lblId.Caption & "'"
End Sub

Private Sub Form_Load()
connectdb

End Sub

Private Sub usernametxt_LostFocus()
If usernametxt.Text = "" Then Exit Sub
If rs.State Then rs.Close
rs.Open "select count(*) from login where username='" & usernametxt.Text & "'", con
If Val(rs(0)) > 0 Then
    MsgBox ("Username already exists")
    usernametxt.Text = ""
End If
End Sub
Sub clear()
usernametxt.Text = ""
addresstxt.Text = ""
dobtxt.Text = ""
qualificationtxt.Text = ""
phonetxt.Text = ""
stafftypetxt.Text = ""
emailtxt.Text = ""
usernametxt.Text = ""
passwordtxt.Text = ""
confirmtxt.Text = ""
End Sub
