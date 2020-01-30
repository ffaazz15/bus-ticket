VERSION 5.00
Begin VB.MDIForm n 
   BackColor       =   &H00FF8080&
   Caption         =   "MDIForm1"
   ClientHeight    =   12075
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17265
   LinkTopic       =   "MDIForm1"
   Picture         =   "proj.frx":0000
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu new 
      Caption         =   "&New Booking"
      Begin VB.Menu munlogin 
         Caption         =   "&login"
      End
      Begin VB.Menu S 
         Caption         =   "&Search Buses"
      End
   End
   Begin VB.Menu a2 
      Caption         =   "Availability "
      Begin VB.Menu s1 
         Caption         =   "&Seat Availability"
      End
   End
   Begin VB.Menu MT1 
      Caption         =   "&My Tickets"
      Begin VB.Menu tb1 
         Caption         =   "Ticket B&ooking"
      End
   End
   Begin VB.Menu H 
      Caption         =   "&Help"
   End
   Begin VB.Menu Ex 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "n"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Boat_Click()
Form6.Show
End Sub

Private Sub Ex_Click()
Me.Hide
End Sub

Private Sub H_Click()
Form5.Show
End Sub

Private Sub munlogin_Click()
Form2.Show
End Sub

Private Sub S_Click()
Form3.Show
End Sub

Private Sub s1_Click()
Form4.Show
End Sub

Private Sub St_Click()

End Sub

Private Sub tb1_Click()
Form6.Show
End Sub
