VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6390
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11760
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnNew 
      Caption         =   "Add"
      Begin VB.Menu smnNewFilm 
         Caption         =   "New Movie"
      End
      Begin VB.Menu smnRelease 
         Caption         =   "Set Shows"
      End
      Begin VB.Menu smnStaff 
         Caption         =   "Staff Registration"
      End
      Begin VB.Menu smnItemSock 
         Caption         =   "Add Item Stock"
      End
      Begin VB.Menu smnAddLine1 
         Caption         =   "-"
      End
      Begin VB.Menu smnRate 
         Caption         =   "Update Screen Rate"
      End
   End
   Begin VB.Menu mnUpdate 
      Caption         =   "View"
      Begin VB.Menu smnViewMovie 
         Caption         =   "View Movies"
      End
      Begin VB.Menu smnStaffs 
         Caption         =   "Staffs"
      End
      Begin VB.Menu smnBookings 
         Caption         =   "Bookings"
      End
      Begin VB.Menu smnSalary 
         Caption         =   "Salary Calculation"
      End
      Begin VB.Menu smnViewSaleas 
         Caption         =   "View Sales"
      End
   End
   Begin VB.Menu mnTicketing 
      Caption         =   "Ticketing"
      Begin VB.Menu smnTktBooking 
         Caption         =   "Ticket Booking"
      End
      Begin VB.Menu tktCancell 
         Caption         =   "Ticket Cancellation"
      End
   End
   Begin VB.Menu mnSales 
      Caption         =   "Sales"
      Begin VB.Menu smnSale 
         Caption         =   "Sale Items"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub smnBookings_Click()
viewBookings.Show
End Sub

Private Sub smnItemSock_Click()
stock.Show
End Sub

Private Sub smnNewFilm_Click()
addfilm.Show
End Sub

Private Sub smnRate_Click()
rate.Show
End Sub

Private Sub smnRelease_Click()
update.Show
End Sub

Private Sub smnSalary_Click()
viewAttendence.Show
End Sub

Private Sub smnSale_Click()
timePassSuply.Show
End Sub

Private Sub smnStaff_Click()
staffreg.Show
End Sub

Private Sub smnStaffs_Click()
viewstaff.Show
End Sub

Private Sub smnTktBooking_Click()
tktBooking.Show
End Sub

Private Sub smnViewMovie_Click()
viewMovies.Show
End Sub

Private Sub smnViewSaleas_Click()
viewStock.Show
End Sub

Private Sub tktCancell_Click()
cancellation.Show
End Sub
