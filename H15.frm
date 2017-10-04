VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   240
      TabIndex        =   18
      Top             =   6480
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "DATE OF JOINING"
      TabPicture(0)   =   "H15.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label10"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label11"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label12"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label13"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Combo1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Combo2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text9"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text10"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text11"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text12"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Command1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      Begin VB.CommandButton Command1 
         Caption         =   "UPDATE"
         Height          =   495
         Left            =   4920
         TabIndex        =   19
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Text12 
         Height          =   405
         Left            =   1920
         TabIndex        =   17
         Top             =   3480
         Width           =   2055
      End
      Begin VB.TextBox Text11 
         Height          =   495
         Left            =   1920
         TabIndex        =   16
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox Text10 
         Height          =   405
         Left            =   1920
         TabIndex        =   15
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox Text9 
         Height          =   405
         Left            =   1920
         TabIndex        =   14
         Top             =   1680
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "H15.frx":001C
         Left            =   8040
         List            =   "H15.frx":0026
         TabIndex        =   9
         Text            =   " "
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   6120
         TabIndex        =   8
         Text            =   " "
         ToolTipText     =   "ENTER RECIEPT NO"
         Top             =   1560
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "H15.frx":0037
         Left            =   6120
         List            =   "H15.frx":0041
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1920
         TabIndex        =   4
         Text            =   " "
         ToolTipText     =   "ENTER DATE OF JOINONG"
         Top             =   4080
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1920
         TabIndex        =   2
         Text            =   " "
         ToolTipText     =   "ENTER REG NO"
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "ROOMNO"
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label12 
         Caption         =   "BLOCJDESC"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "BLOCKNO"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "HOSTELNAME"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "FEE RECIEPT NO"
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "FEES PAID"
         Height          =   375
         Left            =   4560
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "REGNO"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "DATE OF JOINING"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   4200
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WS As Workspace
Dim RS As Recordset
Dim RS1 As Recordset
Dim RS2 As Recordset
Dim rs3 As Recordset
Dim RS4 As Recordset
Dim RS5 As Recordset
Dim RS6 As Recordset
Dim RS7 As Recordset
Dim DB As Database
Dim RS8 As Recordset
Dim rs9 As Recordset
Dim rs10 As Recordset
Dim RS11 As Recordset
Dim s As String
Dim QDF As QueryDef



Private Sub Combo2_Click()
If Combo2.ListIndex = 0 Then
Call ADD1
Else
If Combo2.ListIndex = 1 Then
Call MODI
End If
End If
End Sub

Private Sub Command1_Click()

RS1.Fields(0) = Val(Text1.Text)
RS1.Fields(1) = Text2.Text
RS1.Fields(2) = Text9.Text
RS1.Fields(3) = Text10.Text
RS1.Fields(4) = Text11.Text
RS1.Fields(5) = Val(Text12.Text)
RS1.Fields(6) = Combo1.Text
RS1.Fields(7) = Val(Text3.Text)
RS1.Update
MsgBox "rec updated"
End Sub

Private Sub Form_Load()
Set WS = DBEngine.Workspaces(0)
Set DB = WS.OpenDatabase("C:\program files\nani\stud.mdb", False, False)
Set RS = DB.OpenRecordset("joindetails", dbOpenDynaset)
End Sub
Public Sub Save()
RS.Fields(0) = Val(Text1.Text)
RS.Fields(1) = Text2.Text
RS.Fields(2) = Text9.Text
RS.Fields(3) = Text10.Text
RS.Fields(4) = Text11.Text
RS.Fields(5) = Val(Text12.Text)
RS.Fields(6) = Combo1.Text
RS.Fields(7) = Val(Text3.Text)
RS.Update
End Sub

Public Sub MODI()
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM JOINDETAILS WHERE REGNO=" & InputBox(" ENTER REGNO")
Set RS1 = QDF.OpenRecordset()
If RS1.RecordCount = 0 Then
MsgBox "REGNO does not exist"
Else
Call DIS
Text13.SetFocus
End If
End Sub
Public Sub DIS()
Text1.Text = RS1.Fields(0)
Text2.Text = RS1.Fields(1)
Text3.Text = RS1.Fields(7)
Text9.Text = RS1.Fields(2)
Text10.Text = RS1.Fields(3)
Text11.Text = RS1.Fields(4)
Text12.Text = RS.Fields(5)
Combo1.Text = RS.Fields(6)
End Sub




Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii = 77 Or KeyAscii = 109 Then
RS1.Edit
Text1.SetFocus
End If
End Sub
Public Sub ADD1()
RS.AddNew
Call CLR
Text1.SetFocus
End Sub
Public Sub CLR()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo1.Text = ""
Text3.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
End Sub
Private Sub text3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 97 And KeyAscii <> 83) Or (KeyAscii >= 97 And KeyAscii <= 122 And KeyAscii <> 115) Or (KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8) Or (KeyAscii > 90 And KeyAscii < 97) Then
KeyAscii = 0
MsgBox "INVALID ENTRY"
Else
If KeyAscii = 83 Or KeyAscii = 115 Then
KeyAscii = 0
Call Save
MsgBox "recordsaved"
End If
End If
End Sub
Private Sub text1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 97) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 13) Or (KeyAscii > 90 And KeyAscii < 97) Then
KeyAscii = 0
MsgBox "invalid entry"
Else
If KeyAscii = 13 Then
KeyAscii = 0
Set RS7 = DB.OpenRecordset("ROOM_ALLOC", dbOpenDynaset)
Do While RS7.EOF = False
If RS7(1) = Val(Text1.Text) Then
m = m + 1
End If
RS7.MoveNext
Loop
If m = 0 Then
MsgBox "ROOM ON THIS REGNO NOT BOOKED"
Text1.Text = ""
Text1.SetFocus
Else
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM JOINDETAILS WHERE REGNO=" & Val(Text1.Text)
Set RS11 = QDF.OpenRecordset()
If RS11.RecordCount <> 0 Then
MsgBox "RECORDALREADY EXISTS"
Text1.Text = ""
Text1.SetFocus
Else
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM ROOM_ALLOC WHERE REGNO=" & Val(Text1.Text)
Set rs3 = QDF.OpenRecordset()
Text9.Text = rs3(4)
Text10.Text = rs3(5)
Text11.Text = rs3(6)
Text12.Text = rs3(7)
Text2.SetFocus
End If
End If
End If
End If
End Sub
