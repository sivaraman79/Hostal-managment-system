VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   1305
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu new_det 
      Caption         =   "NEW HOSTEL ENTRY"
      Begin VB.Menu h_entry 
         Caption         =   "HOSTEL ENTRY"
      End
   End
   Begin VB.Menu inmentry 
      Caption         =   "INMATE_INFO"
      Begin VB.Menu in_entry 
         Caption         =   "INMATE_INFORMATION"
      End
   End
   Begin VB.Menu j_details 
      Caption         =   "JOINING_DETAILS"
      Begin VB.Menu join1_detals 
         Caption         =   "DATE OF JOININGS"
      End
      Begin VB.Menu it_issue 
         Caption         =   "ITEMS_ISSUING"
      End
   End
   Begin VB.Menu leave_det 
      Caption         =   "LEAVING_DETAILS"
      Begin VB.Menu del_det 
         Caption         =   "DELETING"
      End
   End
   Begin VB.Menu fdet 
      Caption         =   "FINE_DETAILS"
      Begin VB.Menu fine_det 
         Caption         =   "FINE"
      End
   End
   Begin VB.Menu room_det 
      Caption         =   "ROOM_ALLOC"
      Begin VB.Menu new_det1 
         Caption         =   "NEW ROOM ENTRY"
      End
   End
   Begin VB.Menu rep 
      Caption         =   "REPORTS"
      Begin VB.Menu rep1 
         Caption         =   "Reports"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub del_det_Click()
Form6.Show
End Sub

Private Sub fine_det_Click()
Form8.Show
End Sub

Private Sub h_entry_Click()
Form10.Show
End Sub

Private Sub in_entry_Click()
Form2.Show
End Sub

Private Sub it_issue_Click()
Form5.Show
End Sub

Private Sub join1_detals_Click()
Form4.Show
End Sub

Private Sub new_det1_Click()
Form9.Show
End Sub

Private Sub rep1_Click()
Form7.Show
End Sub
