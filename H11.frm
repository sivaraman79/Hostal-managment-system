VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form9"
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   0
      TabIndex        =   29
      Text            =   " "
      Top             =   5640
      Width           =   375
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "H11.frx":0000
      Height          =   1095
      Left            =   480
      OleObjectBlob   =   "H11.frx":0014
      TabIndex        =   28
      Top             =   5040
      Width           =   7215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\nani\Stud.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox SSTab1 
      Height          =   4695
      Left            =   0
      ScaleHeight     =   4635
      ScaleWidth      =   10755
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   3600
         TabIndex        =   14
         Text            =   " "
         Top             =   3840
         Width           =   1575
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Text            =   " "
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   7800
         TabIndex        =   11
         Text            =   " "
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   7800
         TabIndex        =   10
         Text            =   " "
         Top             =   1800
         Width           =   975
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         ItemData        =   "H11.frx":09E7
         Left            =   1440
         List            =   "H11.frx":09F1
         TabIndex        =   9
         Top             =   3000
         Width           =   1575
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "H11.frx":09FB
         Left            =   4320
         List            =   "H11.frx":0A05
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   7800
         TabIndex        =   7
         Text            =   " "
         ToolTipText     =   "PRESS ""S"" TO SAVE"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   7680
         TabIndex        =   6
         Text            =   " "
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Text            =   " "
         Top             =   4200
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Text            =   "   "
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Text            =   " "
         Top             =   1920
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Text            =   " "
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   4200
         TabIndex        =   1
         Top             =   2640
         Width           =   612
      End
      Begin VB.Label Label11 
         Caption         =   "BRANCH"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "APP_NO"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "NO OF PERSONS"
         Height          =   255
         Left            =   6000
         TabIndex        =   24
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "MAX MEMBERS "
         Height          =   255
         Left            =   6120
         TabIndex        =   23
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "SEX"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "ROOMNO"
         Height          =   375
         Left            =   6120
         TabIndex        =   21
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "BLOCKDESC"
         Height          =   372
         Left            =   6120
         TabIndex        =   20
         Top             =   1200
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "BLOCK"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "HOSTELNAME"
         Height          =   375
         Left            =   0
         TabIndex        =   18
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "NAME"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "REGNO"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "YEAR"
         Height          =   252
         Left            =   3480
         TabIndex        =   15
         Top             =   2640
         Width           =   852
      End
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   372
      Left            =   4200
      TabIndex        =   27
      Top             =   2880
      Width           =   972
   End
End
Attribute VB_Name = "Form9"
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
Dim rs12 As Recordset
Dim RS13 As Recordset
Dim s As String
Dim C As Integer
Dim QDF As QueryDef

Private Sub Combo1_Click()
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT BLOCKTYPE  FROM HOSTEL_INFO WHERE HOSTEL_NAME=" & "'" & Combo1.Text & "'"
Set RS1 = QDF.OpenRecordset()
Combo2.Clear
Do While RS1.EOF = False
Combo2.AddItem RS1(0)
RS1.MoveNext
Loop
End Sub
Private Sub Combo2_Click()
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT BLOCKDESC FROM HOSTEL_INFO WHERE BLOCKTYPE=" & "'" & Combo2.Text & "'" & "AND HOSTEL_NAME=" & "'" & Combo1.Text & "'"
Set RS2 = QDF.OpenRecordset()
Combo3.Clear
Do While RS2.EOF = False
Combo3.AddItem RS2(0)
RS2.MoveNext
Loop
End Sub
Private Sub Combo3_Click()
Dim K As Integer
Dim H As Integer
Dim m As Integer
Dim D As Integer
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM HOSTEL_INFO WHERE BLOCKTYPE=" & "'" & Combo2.Text & "'" & "AND HOSTEL_NAME=" & "'" & Combo1.Text & "'" & "AND BLOCKDESC=" & "'" & Combo3.Text & "'"
Set rs3 = QDF.OpenRecordset()
QDF.SQL = "SELECT * FROM ROOM_ALLOC WHERE BLOCK=" & "'" & Combo2.Text & "'" & "AND HOSTELNAME=" & "'" & Combo1.Text & "'" & "AND BLOCKDESC=" & "'" & Combo3.Text & "'"
Set RS4 = QDF.OpenRecordset()
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "select * from room_info WHERE HOSTEL_NAME=" & "'" & Combo1.Text & "'"
Set RS6 = QDF.OpenRecordset()
Combo4.Clear
Text3.Text = rs3(5)
If RS4.RecordCount = 0 Then
K = rs3(3)
Do While K <> rs3(4)
Combo4.AddItem K
K = K + 1
Loop
Combo4.AddItem K
Else
K = rs3(3)
Do While K <> rs3(4)
Do While RS6.EOF = False
If (RS6(1) = K And RS6(3) = rs3(5)) Then
H = H + 1
End If
RS6.MoveNext
Loop
If H = 0 Then
Combo4.AddItem K
End If
K = K + 1
RS6.MoveFirst
H = 0
Loop
If H = 0 Then
Combo4.AddItem K
End If
End If
Combo4.SetFocus
End Sub
Private Sub Combo4_GOTFOCUS()
Combo4.Text = Combo4.List(C)
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "select * from ROOM_INFO WHERE HOSTEL_NAME=" & "'" & Combo1.Text & "'" & "AND ROOMNO =" & Val(Combo4.Text)
Set RS7 = QDF.OpenRecordset()
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "select application_no,name,sex,roomno,branch from room_alloc where HOSTELNAME=" & "'" & Combo1.Text & "'" & "AND ROOMNO =" & Val(Combo4.Text)
Set RS8 = QDF.OpenRecordset()
Set Data1.Recordset = RS8
If RS8.RecordCount = 0 Then
DBGrid1.Visible = False
Else
DBGrid1.Visible = True
End If
If RS7.RecordCount = 0 Then
Text4.Text = Val(0)
Else
Text4.Text = RS7(3)
End If
End Sub



Private Sub Combo4_KeyUp(KeyCode As Integer, Shift As Integer)
Dim J As Integer
J = Combo4.ListCount
C = C + 1
If C <> J Then
Combo4.Text = Combo4.List(C)
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "select * from ROOM_INFO WHERE HOSTEL_NAME=" & "'" & Combo1.Text & "'" & "AND ROOMNO =" & Val(Combo4.Text)
Set RS7 = QDF.OpenRecordset()
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "select application_no,name,sex,roomno,branch from room_alloc where HOSTELNAME=" & "'" & Combo1.Text & "'" & "AND ROOMNO =" & Val(Combo4.Text)
Set RS8 = QDF.OpenRecordset()
Set Data1.Recordset = RS8
If RS8.RecordCount = 0 Then
DBGrid1.Visible = False
Else
DBGrid1.Visible = True
End If
If RS7.RecordCount = 0 Then
Text4.Text = Val(0)
Else
Text4.Text = RS7(3)
End If
Else
Combo4.Text = Combo4.List(C - 1)
MsgBox "NO RECORDS"
C = -1
End If
End Sub

Private Sub Combo5_CLICK()
If Combo5.ListIndex = 0 Then
Call ADD1
C = 0
Else
If Combo5.ListIndex = 1 Then
Call CLR
Call DEL
End If
End If
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
Dim t As Integer
Dim J As String
Dim m As Integer
If KeyAscii = 83 Or KeyAscii = 115 Then
KeyAscii = 0
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "select * from room_alloc where hostelname=" & "'" & Combo1.Text & "'"
Set rs10 = QDF.OpenRecordset()
If rs10.RecordCount = 0 Then
Text6.Text = Left(Combo1.Text, 3) + LTrim(Str(100))
Else
rs10.MoveLast
t = Len(rs10(9))
m = Right(rs10(9), t - 3)
t = m + 1
J = Str(t)
Text6.Text = Left(Combo1.Text, 3) + LTrim(J)
End If
Call Save
MsgBox "RECORDSAVED"
Combo5.SetFocus
End If
End Sub



Private Sub Combo7_Click()
Dim i As Integer
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "select * from branch_info where bname=" & "'" & Combo7.Text & "'"
Set rs12 = QDF.OpenRecordset()
i = 1
Do While i <= rs12(2)
Combo8.AddItem i
i = i + 1
Loop
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim m As Integer
Set WS = DBEngine.Workspaces(0)
Set DB = WS.OpenDatabase("C:\program files\nani\stud.mdb", False, False)
Set RS = DB.OpenRecordset("hostel_info", dbOpenDynaset)
Set RS5 = DB.OpenRecordset("ROOM_ALLOC", dbOpenDynaset)
Do While RS.EOF = False
If Combo1.ListCount = 0 Then
Combo1.AddItem RS(0)
Else
K = Combo1.ListCount
Do While K <> i
If RS(0) = Combo1.List(i) Then
m = m + 1
End If
i = i + 1
Loop
If m = 0 Then
Combo1.AddItem RS(0)
End If
End If
i = 0
m = 0
RS.MoveNext
Loop
Set RS8 = DB.OpenRecordset("BRANCH_INFO", dbOpenDynaset)
Combo7.Clear
Do While RS8.EOF = False
Combo7.AddItem RS8(1)
RS8.MoveNext
Loop
End Sub
Public Sub ADD1()
DBGrid1.Visible = False
Call CLR
RS5.AddNew
Text5.SetFocus
End Sub
Public Sub CLR()
Text1.Text = ""
Text5.Text = ""
Text6.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo6.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Combo4.Text = ""
Combo7.Text = ""
Combo8.Clear
End Sub
Public Sub Save()
RS5.Fields(0) = Val(Text5.Text)
RS5.Fields(1) = Val(Text1.Text)
RS5.Fields(2) = Text2.Text
RS5.Fields(3) = Combo6.Text
RS5.Fields(4) = Combo1.Text
RS5.Fields(5) = Combo2.Text
RS5.Fields(6) = Combo3.Text
RS5.Fields(7) = Combo4.Text
RS5.Fields(8) = Combo7.Text
RS5.Fields(9) = Text6.Text
RS5.Fields(10) = Combo8.Text
RS5.Update
If RS7.RecordCount = 0 Then
RS7.AddNew
RS7.Fields(0) = Combo1.Text
RS7.Fields(1) = Val(Combo4.Text)
RS7.Fields(2) = Combo3.Text
RS7.Fields(3) = Val(1)
RS7.Update
Else
RS7.Edit
RS7.Fields(3) = RS7.Fields(3) + 1
RS7.Update
End If
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 97) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8) Or (KeyAscii > 90 And KeyAscii < 97) Then
KeyAscii = 0
MsgBox "invalid entry"
Else
If KeyAscii = 13 Then
KeyAscii = 0
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM ROOM_ALLOC WHERE REGNO=" & Val(Text1.Text)
Set RS11 = QDF.OpenRecordset()
If RS11.RecordCount <> 0 Then
MsgBox "RECORDALREADY EXISTS"
Text1.Text = ""
Text1.SetFocus
Else
Text2.SetFocus
End If
End If
End If
End Sub

Private Sub TEXT7_KeyPress(KeyAscii As Integer)
If KeyAscii = 68 Or KeyAscii = 100 Then
KeyAscii = 0
If MsgBox("DO U WANT TO DELETE THE REC", vbYesNo) = vbYes Then
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "select * from room_info WHERE HOSTEL_NAME=" & "'" & Combo1.Text & "'" & "AND ROOMNO =" & Val(Combo4.Text)
Set RS6 = QDF.OpenRecordset()
RS6.Edit
RS6(3) = Val(RS6(3) - 1)
RS6.Update
If RS6(3) <= 0 Then
RS6.Delete
End If
RS13.Delete
Call CLR
End If
End If
End Sub
Private Sub text2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox "invalid entry"
Else
If (KeyAscii >= 97 And KeyAscii >= 122) Then
KeyAscii = 0
MsgBox "enter only capitals"
End If
End If
End Sub

Private Sub text5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 97) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8) Or (KeyAscii > 90 And KeyAscii < 97) Then
KeyAscii = 0
MsgBox "invalid entry"
Else
If KeyAscii = 13 Then
KeyAscii = 0
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM ROOM_ALLOC WHERE APPLICATION_NO=" & Val(Text5.Text)
Set RS11 = QDF.OpenRecordset()
If RS11.RecordCount <> 0 Then
MsgBox "RECORDALREADY EXISTS"
Text5.Text = ""
Text5.SetFocus
Else
Text1.SetFocus
End If
End If
End If
End Sub



Public Sub DEL()
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM ROOM_ALLOC WHERE REGNO=" & InputBox(" ENTER REG NO")
On Error GoTo 2
Set RS13 = QDF.OpenRecordset()
If RS13.RecordCount = 0 Then
MsgBox "REGNO does not exist"
Else
Call DIS
Text7.SetFocus
End If
2  End Sub

'Public Sub MODI()
'Set QDF = DB.CreateQueryDef("")
'QDF.SQL = "SELECT * FROM  ROOM_ALLOC WHERE REGNO=" & "'" & InputBox(" ENTER HOSTEL NAME") & "'"
'Set RS13 = QDF.OpenRecordset()
'If RS13.RecordCount = 0 Then
'MsgBox "REGNO does not exist"
'Else
'Call DIS
'Text7.SetFocus
'End If
'End Sub

Public Sub DIS()
Text5.Text = RS13(0)
Text1.Text = RS13(1)
Text2.Text = RS13(2)
Combo6.Text = RS13(3)
Combo1.Text = RS13(4)
Combo2.Text = RS13(5)
Combo3.Text = RS13(6)
Combo4.Text = RS13(7)
Combo7.Text = RS13(8)
Text6.Text = RS13(9)
Combo8.Text = RS13(10)
End Sub

'Private Sub Command1_Click()
'RS13.Fields(0) = Val(Text5.Text)
'RS13.Fields(1) = Val(Text1.Text)
'RS13.Fields(2) = Text2.Text
'RS13.Fields(3) = Combo6.Text
'RS13.Fields(4) = Combo1.Text
'RS13.Fields(5) = Combo2.Text
'RS13.Fields(6) = Combo3.Text
'RS13.Fields(7) = Combo4.Text
'RS13.Fields(8) = Combo7.Text
'RS13.Fields(9) = Text6.Text
'RS13.Fields(10) = Combo8.Text
'RS13.Update
'MsgBox "rec updated"
'End Sub
