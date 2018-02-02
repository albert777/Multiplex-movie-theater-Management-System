VERSION 5.00
Begin VB.Form rate 
   Caption         =   "Form3"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12030
   LinkTopic       =   "Form3"
   ScaleHeight     =   6435
   ScaleWidth      =   12030
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   4440
      TabIndex        =   5
      Top             =   2520
      Width           =   3975
      Begin VB.CheckBox Check4 
         Caption         =   "4"
         Height          =   615
         Left            =   3000
         TabIndex        =   9
         Top             =   120
         Width           =   735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "3"
         Height          =   615
         Left            =   2040
         TabIndex        =   8
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Caption         =   "2"
         Height          =   615
         Left            =   1080
         TabIndex        =   7
         Top             =   120
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "1"
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   4560
      TabIndex        =   4
      Top             =   3720
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4560
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Amount"
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Select  Show"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Select Screen"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   1560
      Width           =   2655
   End
End
Attribute VB_Name = "rate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
