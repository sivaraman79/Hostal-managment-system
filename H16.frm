VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   10186
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "H16.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(8)=   "LIST1"
      Tab(0).Control(9)=   "Text1"
      Tab(0).Control(10)=   "Text2"
      Tab(0).Control(11)=   "Text3"
      Tab(0).Control(12)=   "Text4"
      Tab(0).Control(13)=   "Text5"
      Tab(0).Control(14)=   "Text6"
      Tab(0).Control(15)=   "Text7"
      Tab(0).Control(16)=   "Text8"
      Tab(0).Control(17)=   "Combo1"
      Tab(0).Control(18)=   "Text9"
      Tab(0).Control(19)=   "Command1"
      Tab(0).Control(20)=   "Command2"
      Tab(0).Control(21)=   "Command3"
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "H16.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "H16.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.ListBox LIST1 
         Height          =   3375
         ItemData        =   "H16.frx":0054
         Left            =   -74520
         List            =   "H16.frx":0056
         TabIndex        =   15
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   -70560
         TabIndex        =   14
         Text            =   " "
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   -70560
         TabIndex        =   13
         Text            =   " "
         Top             =   1920
         Width           =   5415
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   -70560
         TabIndex        =   12
         Text            =   " "
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   -70560
         TabIndex        =   11
         Text            =   " "
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   -70560
         TabIndex        =   10
         Text            =   " "
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   -70560
         TabIndex        =   9
         Text            =   " "
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   -70560
         TabIndex        =   8
         Text            =   " "
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   -67560
         TabIndex        =   7
         Text            =   " "
         Top             =   3600
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "H16.frx":0058
         Left            =   -68160
         List            =   "H16.frx":0062
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   -74880
         TabIndex        =   5
         Text            =   "Text9"
         Top             =   5760
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "NEXT"
         Height          =   255
         Left            =   -68880
         TabIndex        =   4
         Top             =   4920
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "PREVIOUS"
         Height          =   255
         Left            =   -66960
         TabIndex        =   3
         Top             =   4920
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "UPDATE"
         Height          =   255
         Left            =   -68040
         TabIndex        =   2
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "REGNO"
         Height          =   255
         Left            =   -71760
         TabIndex        =   23
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "NAME"
         Height          =   255
         Left            =   -71760
         TabIndex        =   22
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "HOSTELNAME"
         Height          =   375
         Left            =   -71760
         TabIndex        =   21
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "BLOCKNAME"
         Height          =   255
         Left            =   -71760
         TabIndex        =   20
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "BLOCKDESC"
         Height          =   375
         Left            =   -71760
         TabIndex        =   19
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "ROOMNO"
         Height          =   375
         Left            =   -71760
         TabIndex        =   18
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "ITEMNAME"
         Height          =   495
         Left            =   -71760
         TabIndex        =   17
         Top             =   4920
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "ITEMCODE"
         Height          =   375
         Left            =   -68640
         TabIndex        =   16
         Top             =   3600
         Width           =   855
      End
   End
   Begin VB.Label Label9 
      Caption         =   "FOR MODIFICATION SELECT THE RECORD AND PRESS M AND AFTER MODIFICATION  PRESS UPDATE"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   6840
      Width           =   8535
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WS As Workspace
Dim RS As Recordset
Dim DB As Database
Dim RS1 As Recordset
Dim RS2 As Recordset
Dim rs3 As Recordset
Dim RS4 As Recordset
Dim QDF As QueryDef
Dim RS7  As Recordset
Dim RS8 As Recordset

Public Sub ADD1()
Call CLR
Text1.SetFocus
End Sub



Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
Call ADD1
Else
If Combo1.ListIndex = 1 Then
Call MODI
End If
End If
End Sub

Private Sub Form_Load()
Set WS = DBEngine.Workspaces(0)
Set DB = WS.OpenDatabase("C:\program files\nani\stud.mdb", False, False)
Set RS = DB.OpenRecordset("ITEM_ALLOC", dbOpenDynaset)
Set rs3 = DB.OpenRecordset("ITEM_DESC", dbOpenDynaset)
Do While rs3.EOF = False
List1.AddItem rs3(1)
rs3.MoveNext
Loop
End Sub
Public Sub CLR()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
End Sub

Private Sub list1_Click()
Text7.Text = ""
Text8.Text = ""
Text7.SetFocus
Text7.Text = List1.List(List1.ListIndex)
Text8.SetFocus
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
Dim m As Integer
If (KeyAscii >= 65 And KeyAscii <= 97) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8) Or (KeyAscii > 90 And KeyAscii < 97) Then
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
QDF.SQL = "SELECT * FROM ITEM_ALLOC WHERE REGNO=" & Val(Text1.Text)
Set RS11 = QDF.OpenRecordset()
If RS11.RecordCount <> 0 Then
MsgBox "RECORDALREADY EXISTS"
Text1.Text = ""
Text1.SetFocus
Else
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM ROOM_ALLOC WHERE REGNO=" & Val(Text1.Text)
Set rs3 = QDF.OpenRecordset()
Text2.Text = rs3(2)
Text3.Text = rs3(4)
Text4.Text = rs3(5)
Text5.Text = rs3(6)
Text6.Text = rs3(7)
Text7.SetFocus
End If
End If
End If
End If
End Sub

Public Sub Save()
RS.AddNew
RS.Fields(0) = Val(Text1.Text)
RS.Fields(1) = Text2.Text
RS.Fields(2) = Text3.Text
RS.Fields(3) = Text4.Text
RS.Fields(4) = Text5.Text
RS.Fields(5) = Val(Text6.Text)
RS.Fields(6) = Text7.Text
RS.Fields(7) = Text8.Text
RS.Update
End Sub

Private Sub TEXT8_KeyPress(KeyAscii As Integer)
If KeyAscii = 83 Or KeyAscii = 115 Then
KeyAscii = 0
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM ITEM_ALLOC WHERE ITEMCODE=" & "'" & Text8.Text & "'"
Set RS4 = QDF.OpenRecordset()
If RS4.RecordCount <> 0 Then
MsgBox "ITEMCODE ALREADYEXISTS"
Text8.Text = ""
Text8.SetFocus
Else
Call Save
MsgBox "recordsaved"
End If
End If
End Sub
Public Sub MODI()
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM ITEM_ALLOC WHERE REGNO=" & InputBox(" ENTER REGNO")
Set RS7 = QDF.OpenRecordset()
If RS7.RecordCount = 0 Then
MsgBox "REGNO does not exist"
Else
Call DIS
Text9.SetFocus
End If
End Sub

Public Sub DIS()
Text1.Text = RS7(0)
Text2.Text = RS7(1)
Text3.Text = RS7(2)
Text4.Text = RS7(3)
Text5.Text = RS7(4)
Text6.Text = RS7(5)
Text7.Text = RS7(6)
Text8.Text = RS7(7)
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 77 Or KeyAscii = 109 Then
RS7.Edit
Text1.SetFocus
End If
End Sub

Private Sub Command1_Click()
RS7.MoveNext
If RS7.EOF = True Then
MsgBox "NO RECORDS"
RS7.MoveLast
Call DIS
Text9.SetFocus
Else
Call DIS
Text9.SetFocus
End If
End Sub
Private Sub Command2_Click()
RS7.MovePrevious
If RS7.BOF = True Then
MsgBox "NO RECORDS"
RS7.MoveFirst
Call DIS
Text9.SetFocus
Else
Call DIS
Text9.SetFocus
End If
End Sub
Private Sub Command3_Click()
RS7.Fields(0) = Val(Text1.Text)
RS7.Fields(1) = Text2.Text
RS7.Fields(2) = Text3.Text
RS7.Fields(3) = Text4.Text
RS7.Fields(4) = Text5.Text
RS7.Fields(5) = Val(Text6.Text)
RS7.Fields(6) = Text7.Text
RS7.Fields(7) = Text8.Text
RS7.Update
MsgBox "rec updated"
End Sub
