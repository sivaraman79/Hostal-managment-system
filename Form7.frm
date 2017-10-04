VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "TO MAIN MENU"
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stud Info"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Text            =   " "
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Item Alloc"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fine Detail"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hostel Info"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Regno"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Regno "
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Regno "
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DataReport1.Show
End Sub

Private Sub Command2_Click()
Text1.Visible = True
Label1.Visible = True
Text1.SetFocus
End Sub

Private Sub Command3_Click()
Text2.Visible = True
Label2.Visible = True
Text2.SetFocus
End Sub

Private Sub Command4_Click()
Text3.Visible = True
Label3.Visible = True
Text3.SetFocus
End Sub

Private Sub Command5_Click()
MDIForm1.Show
Form7.Hide
End Sub

Private Sub text1_lostfocus()
Text1.Visible = False
Label1.Visible = False
DataEnvironment1.rsCommand2.Open "select * from fine_details where regno=" & Val(Text1.Text)
DataEnvironment1.rsCommand2.Requery
If DataEnvironment1.rsCommand2.BOF = False And DataEnvironment1.rsCommand2.EOF = False Then
DataReport2.Refresh
DataReport2.WindowState = 2
DataReport2.Show
Else
 MsgBox "There are no Records"
End If
DataEnvironment1.rsCommand2.Close
End Sub
Private Sub text2_lostfocus()
Text2.Visible = False
Label2.Visible = False
DataEnvironment1.rsCommand4.Open "select * from item_alloc where regno=" & Val(Text2.Text)
DataEnvironment1.rsCommand4.Requery
If DataEnvironment1.rsCommand4.BOF = False And DataEnvironment1.rsCommand4.EOF = False Then
DataReport4.Refresh
DataReport4.WindowState = 2
DataReport4.Show
Else
 MsgBox "There are no Records"
End If
DataEnvironment1.rsCommand4.Close
End Sub
Private Sub text3_lostfocus()
Text3.Visible = False
Label3.Visible = False
DataEnvironment1.rsCommand3.Open "select * from stud_info where reg_no=" & Val(Text3.Text)
DataEnvironment1.rsCommand3.Requery
If DataEnvironment1.rsCommand3.BOF = False And DataEnvironment1.rsCommand3.EOF = False Then
DataReport3.Refresh
DataReport3.WindowState = 2
DataReport3.Show
Else
 MsgBox "There are no Records"
End If
DataEnvironment1.rsCommand3.Close
End Sub

