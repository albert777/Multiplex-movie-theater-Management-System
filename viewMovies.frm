VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form viewMovies 
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "viewMovies.frx":0000
   ScaleHeight     =   5550
   ScaleWidth      =   11565
   WindowState     =   2  'Maximized
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
      ItemData        =   "viewMovies.frx":0342
      Left            =   1920
      List            =   "viewMovies.frx":0352
      TabIndex        =   3
      Text            =   "1"
      Top             =   1800
      Visible         =   0   'False
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
      ItemData        =   "viewMovies.frx":0362
      Left            =   6720
      List            =   "viewMovies.frx":0372
      TabIndex        =   2
      Text            =   "1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   9480
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   11295
      _ExtentX        =   19923
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
   Begin VB.Label Label6 
      Caption         =   "Show"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Screen"
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "viewMovies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub comboscreen_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub comboShow_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_Load()
connectdb
If rs.State Then rs.Close
rs.Open "select * from addfilm ", con
Set MSHFlexGrid1.DataSource = rs
End Sub
