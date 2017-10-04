VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   3195
   ClientLeft      =   -2055
   ClientTop       =   675
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15266
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "H18.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "DBGrid1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "DBGrid2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text11"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text9"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Combo1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Data2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text8"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text7"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text6"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text5"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text4"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text3"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text10"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Data1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Command1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "H18.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(1)=   "Label12"
      Tab(1).Control(2)=   "Text12"
      Tab(1).Control(3)=   "DBGrid3"
      Tab(1).Control(4)=   "Data3"
      Tab(1).Control(5)=   "Text13"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   -65760
         TabIndex        =   30
         Text            =   " "
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   "C:\Program Files\nani\Stud.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -73440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5760
         Visible         =   0   'False
         Width           =   3780
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "H18.frx":0038
         Height          =   1575
         Left            =   -75000
         OleObjectBlob   =   "H18.frx":004C
         TabIndex        =   28
         Top             =   2040
         Width           =   11775
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   -73560
         TabIndex        =   26
         Text            =   " "
         Top             =   840
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "UPDATE"
         Height          =   375
         Left            =   7680
         TabIndex        =   25
         Top             =   2400
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Program Files\nani\Stud.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   8040
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   1080
         TabIndex        =   24
         Text            =   " "
         Top             =   8400
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1920
         TabIndex        =   13
         Text            =   " "
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Text            =   " "
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1920
         TabIndex        =   11
         Text            =   " "
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Text            =   " "
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Text            =   " "
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Text            =   " "
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   5160
         TabIndex        =   7
         Text            =   " "
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   5160
         TabIndex        =   6
         Text            =   " "
         Top             =   4560
         Width           =   2775
      End
      Begin VB.Data Data2 
         Connect         =   "Access"
         DatabaseName    =   "C:\Program Files\nani\Stud.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1920
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "H18.frx":0A1F
         Left            =   5520
         List            =   "H18.frx":0A2C
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Text            =   " "
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Text            =   " "
         Top             =   360
         Width           =   1695
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "H18.frx":0A45
         Height          =   855
         Left            =   4320
         OleObjectBlob   =   "H18.frx":0A59
         TabIndex        =   4
         Top             =   3000
         Visible         =   0   'False
         Width           =   3375
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "H18.frx":142C
         Height          =   1215
         Left            =   0
         OleObjectBlob   =   "H18.frx":1440
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   5160
         Visible         =   0   'False
         Width           =   10935
      End
      Begin VB.Label Label12 
         Caption         =   "TOTAL"
         DataSource      =   "T"
         Height          =   255
         Left            =   -67080
         TabIndex        =   29
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "REGNO"
         Height          =   255
         Left            =   -74760
         TabIndex        =   27
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "REGNO"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "STUDENTNAME"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "HOSTELNAME"
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "BLOCKTYPE"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "BLOCKDESC"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "FINEDESC"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "FINEAMOUNT"
         Height          =   255
         Left            =   3720
         TabIndex        =   17
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "DATE"
         Height          =   255
         Left            =   3720
         TabIndex        =   16
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "ROOMNO"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "FINENO"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WS As Workspace
Dim RS As Recordset
Dim DB As Database
Dim RS1 As Recordset
Dim RS2 As Recordset
Dim RS4 As Recordset
Dim rs3 As Recordset
Dim RS7 As Recordset
Dim rs10 As Recordset
Dim rs15 As Recordset
Dim QDF As QueryDef

Private Sub DBGrid1_Click()
Dim H As Integer
Dim K As Integer
H = DBGrid1.Row
DBGrid1.Col = 0
Text11.Text = DBGrid1.Text
K = K + 1
DBGrid1.Col = K
Text1.Text = DBGrid1.Text
K = K + 1
DBGrid1.Col = K
Text2.Text = DBGrid1.Text
K = K + 1
DBGrid1.Col = K
Text3.Text = DBGrid1.Text
K = K + 1
DBGrid1.Col = K
Text4.Text = DBGrid1.Text
K = K + 1
DBGrid1.Col = K
Text5.Text = DBGrid1.Text
K = K + 1
DBGrid1.Col = K
Text9.Text = DBGrid1.Text
K = K + 1
DBGrid1.Col = K
Text6.Text = DBGrid1.Text
K = K + 1
DBGrid1.Col = K
Text7.Text = DBGrid1.Text
K = K + 1
DBGrid1.Col = K
Text8.Text = DBGrid1.Text
Text10.SetFocus
End Sub

'Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
'Dim H As Integer
'If KeyAscii = 13 Then
'KeyAscii = 0
'H = DBGrid1.Row
'MsgBox H
'End If
'End Sub

Private Sub DBGrid2_Click()
Dim i As Integer
DBGrid2.Row = 0
DBGrid2.Col = 0
Text2.Text = DBGrid2.Text
DBGrid2.Col = 1
Text3.Text = DBGrid2.Text
DBGrid2.Col = 2
Text4.Text = DBGrid2.Text
DBGrid2.Col = 3
Text5.Text = DBGrid2.Text
DBGrid2.Col = 4
Text9.Text = DBGrid2.Text
DBGrid2.Visible = False

End Sub

Private Sub Form_Load()
Set WS = DBEngine.Workspaces(0)
Set DB = WS.OpenDatabase("C:\program files\nani\stud.mdb", False, False)
Set RS = DB.OpenRecordset("FINE_DETAILS", dbOpenDynaset)
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
Dim m As Integer
If (KeyAscii >= 65 And KeyAscii <= 97) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8) Or (KeyAscii > 90 And KeyAscii < 97) Then
KeyAscii = 0
MsgBox "invalid entry"
Else
If KeyAscii = 13 Then
KeyAscii = 0
Set RS2 = DB.OpenRecordset("ROOM_ALLOC", dbOpenDynaset)
Do While RS2.EOF = False
If RS2(1) = Val(Text1.Text) Then
m = m + 1
End If
RS2.MoveNext
Loop
If m = 0 Then
MsgBox "ROOM ON THIS REGNO NOT BOOKED"
Text1.Text = ""
Text1.SetFocus
Else
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT NAME,HOSTELNAME,BLOCK,BLOCKDESC,ROOMNO FROM ROOM_ALLOC WHERE REGNO=" & Val(Text1.Text)
Set RS4 = QDF.OpenRecordset()
DBGrid2.Visible = True
Set Data2.Recordset = RS4
End If
End If
End If
m = 0
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
Dim J As Integer
If KeyAscii = 13 Then
KRYASCII = 0
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT FNO,REGNO,STUDENTNAME,HOSTELNAME,BLOCKTYPE,BLOCKDESC,ROOMNO,FINEDESC,FINEAMOUNT,DAT FROM FINE_DETAILS WHERE REGNO=" & Val(Text12.Text)
Set rs3 = QDF.OpenRecordset()
Do While rs3.EOF = False
J = J + rs3(8)
rs3.MoveNext
Loop
If rs3.RecordCount <> 0 Then
DBGrid3.Visible = True
Set Data3.Recordset = rs3
Else
If rs3.RecordCount = 0 Then
MsgBox "RECORD DOESNOTEXIST"
Text12.Text = ""
Text12.SetFocus
End If
End If
End If
Text13.Text = Val(J)
End Sub


Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Text7.SetFocus
Else
If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8) Or (KeyAscii > 90 And KeyAscii < 97) Then
KeyAscii = 0
MsgBox "invalid entry"
End If
End If
End Sub



Private Sub TEXT7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Text8.SetFocus
Else
If (KeyAscii >= 65 And KeyAscii <= 97 And KeyAscii <> 83) Or (KeyAscii >= 97 And KeyAscii <= 122 And KeyAscii <> 115) Or (KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8) Or (KeyAscii > 90 And KeyAscii < 97) Then
KeyAscii = 0
MsgBox "INVALID ENTRY"
End If
End If
End Sub

Private Sub TEXT8_KeyPress(KeyAscii As Integer)
If KeyAscii = 83 Or KeyAscii = 115 Then
KeyAscii = 0
Call Save
MsgBox "RECORD SAVED"
End If
DBGrid1.Visible = False
End Sub

Private Sub Text9_Change()
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT FNO,REGNO,STUDENTNAME,HOSTELNAME,BLOCKTYPE,BLOCKDESC,ROOMNO,FINEDESC,FINEAMOUNT,DAT FROM FINE_DETAILS WHERE REGNO=" & Val(Text1.Text)
Set rs3 = QDF.OpenRecordset()
If rs3.RecordCount <> 0 Then
DBGrid1.Visible = True
Set Data1.Recordset = rs3
End If
Text6.SetFocus
End Sub
Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
DBGrid1.Visible = False
Call CLR
Call ADD1
Else
If Combo1.ListIndex = 1 Then
DBGrid1.Visible = False
Call CLR
Call DEL
Else
DBGrid1.Visible = False
Call CLR
Call MODI
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
RS.Fields(5) = Val(Text9.Text)
RS.Fields(6) = Text6.Text
RS.Fields(7) = Val(Text7.Text)
RS.Fields(8) = Text8.Text
RS.Fields(9) = Text11.Text
RS.Update
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
Text9.Text = ""
Text11.Text = ""
End Sub
Public Sub ADD1()
Call CLR
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "select * from FINE_DETAILS"
Set rs10 = QDF.OpenRecordset()
MsgBox rs10.RecordCount
If rs10.RecordCount = 0 Then
Text11.Text = "F" + LTrim(Str(100))
Else
rs10.MoveLast
t = Len(rs10(9))
m = Right(rs10(9), t - 1)
t = m + 1
J = Str(t)
Text11.Text = "F" + LTrim(J)
Text1.SetFocus
End If
End Sub
Public Sub MODI()
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM FINE_DETAILS WHERE REGNO=" & InputBox(" ENTER REGNO")
On Error GoTo 2
Set rs15 = QDF.OpenRecordset()
If rs15.RecordCount = 0 Then
MsgBox "REGNO does not exist"
Else
Call DIS
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT FNO,REGNO,STUDENTNAME,HOSTELNAME,BLOCKTYPE,BLOCKDESC,ROOMNO,FINEDESC,FINEAMOUNT,DAT FROM FINE_DETAILS WHERE REGNO=" & Val(Text1.Text)
Set rs3 = QDF.OpenRecordset()
If rs3.RecordCount <> 0 Then
DBGrid1.Visible = True
Set Data1.Recordset = rs3
Text10.SetFocus
Text10.SetFocus
End If
End If

2 End Sub

Public Sub DIS()
Text1.Text = rs15(0)
Text2.Text = rs15(1)
Text3.Text = rs15(2)
Text4.Text = rs15(3)
Text5.Text = rs15(4)
Text9.Text = rs15(5)
Text6.Text = rs15(6)
Text7.Text = rs15(7)
Text8.Text = rs15(8)
Text11.Text = rs15(9)
End Sub
Public Sub DEL()
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM FINE_DETAILS WHERE REGNO=" & InputBox(" ENTER REGNO OR IF U WANT TO CANCEL PRESS C")
On Error GoTo 2
Set rs15 = QDF.OpenRecordset()
If rs15.RecordCount = 0 Then
MsgBox "regno does not exist"
Else
Call DIS
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT FNO,REGNO,STUDENTNAME,HOSTELNAME,BLOCKTYPE,BLOCKDESC,ROOMNO,FINEDESC,FINEAMOUNT,DAT FROM FINE_DETAILS WHERE REGNO=" & Val(Text1.Text)
Set rs3 = QDF.OpenRecordset()
If rs3.RecordCount <> 0 Then
DBGrid1.Visible = True
Set Data1.Recordset = rs3
Text10.SetFocus
End If
End If

2 End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 68 Or KeyAscii = 100 Then
KeyAscii = 0
If MsgBox("DO U WANT TO DELETE THE REC", vbYesNo) = vbYes Then
rs3.Delete
End If
Else
If KeyAscii = 77 Or KeyAscii = 109 Then
Command1.Visible = True
rs3.Edit
Text1.SetFocus
End If
End If
End Sub
Private Sub Command1_Click()
rs3.Fields(0) = Val(Text1.Text)
rs3.Fields(1) = Text2.Text
rs3.Fields(2) = Text3.Text
rs3.Fields(3) = Text4.Text
rs3.Fields(4) = Text5.Text
rs3.Fields(5) = Val(Text9.Text)
rs3.Fields(6) = Text6.Text
rs3.Fields(7) = Val(Text7.Text)
rs3.Fields(8) = Text8.Text
rs3.Fields(9) = Text11.Text
rs3.Update
MsgBox "rec updated"
Command1.Visible = False

End Sub
