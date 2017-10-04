VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form10"
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox SSTab1 
      Height          =   6135
      Left            =   0
      ScaleHeight     =   6075
      ScaleWidth      =   10515
      TabIndex        =   1
      Top             =   0
      Width           =   10572
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Top             =   5520
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Text            =   " "
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2160
         TabIndex        =   11
         Text            =   " "
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2160
         TabIndex        =   10
         Text            =   " "
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Text            =   " "
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Text            =   "   "
         Top             =   840
         Width           =   1815
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Program Files\Stud.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   -72840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "hostel_info"
         Top             =   5520
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         Height          =   288
         ItemData        =   "H12.frx":0000
         Left            =   4680
         List            =   "H12.frx":000D
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "NEXT"
         Height          =   495
         Left            =   1200
         TabIndex        =   6
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "PREVIOUS"
         Height          =   495
         Left            =   2640
         TabIndex        =   5
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "UPDATE"
         Height          =   495
         Left            =   4080
         TabIndex        =   4
         Top             =   3360
         Width           =   975
      End
      Begin VB.PictureBox DBGrid1 
         Height          =   3732
         Left            =   -74160
         ScaleHeight     =   3675
         ScaleWidth      =   9315
         TabIndex        =   3
         Top             =   600
         Width           =   9372
      End
      Begin VB.TextBox Text7 
         Height          =   288
         Left            =   2160
         TabIndex        =   2
         Text            =   " "
         Top             =   2040
         Width           =   1332
      End
      Begin VB.Label Label5 
         Caption         =   "roomendno"
         Height          =   252
         Left            =   1080
         TabIndex        =   20
         Top             =   2760
         Width           =   1092
      End
      Begin VB.Label Label4 
         Caption         =   "roomstartno"
         Height          =   372
         Left            =   1080
         TabIndex        =   19
         Top             =   2400
         Width           =   1092
      End
      Begin VB.Label Label3 
         Caption         =   "blockdesc"
         Height          =   372
         Left            =   1080
         TabIndex        =   18
         Top             =   1680
         Width           =   1092
      End
      Begin VB.Label Label2 
         Caption         =   "blocktype"
         Height          =   252
         Left            =   1080
         TabIndex        =   17
         Top             =   1320
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "hostelname"
         Height          =   372
         Left            =   1080
         TabIndex        =   16
         Top             =   840
         Width           =   852
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   10560
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label6 
         Caption         =   " FOR DELETION SELECT THE RECORD AND PRESS D"
         Height          =   255
         Left            =   600
         TabIndex        =   15
         Top             =   4440
         Width           =   4335
      End
      Begin VB.Label Label7 
         Caption         =   "FOR MODIFICATION SELECT THE REC AND PRESS M,AFTER MODIFICATION PRESS UPDATE"
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   4920
         Width           =   7335
      End
      Begin VB.Line Line2 
         X1              =   480
         X2              =   480
         Y1              =   4320
         Y2              =   5400
      End
      Begin VB.Line Line3 
         X1              =   480
         X2              =   9480
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line4 
         X1              =   9480
         X2              =   9480
         Y1              =   4320
         Y2              =   5400
      End
      Begin VB.Line Line5 
         X1              =   480
         X2              =   9480
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Label Label8 
         Caption         =   "max persons"
         Height          =   252
         Left            =   1080
         TabIndex        =   13
         Top             =   2040
         Width           =   1332
      End
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   " "
      Top             =   5640
      Width           =   255
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WS As Workspace
Dim RS As Recordset
Dim DB  As Database
Dim RS1 As Recordset
Dim QDF As QueryDef
Public Sub ADD1()
Call CLR
RS.AddNew
Text1.SetFocus
End Sub
Public Sub CLR()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text7.Text = ""
End Sub
Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
Call ADD1
Else
If Combo1.ListIndex = 1 Then
Call DEL
Else
If Combo1.ListIndex = 2 Then
Call MODI
End If
End If
End If
End Sub
Private Sub Command1_Click()
RS1.MoveNext
If RS1.EOF = True Then
MsgBox "NO RECORDS"
RS1.MoveLast
Call DIS
Text8.SetFocus
Else
Call DIS
Text8.SetFocus
End If
End Sub
Private Sub Command2_Click()
RS1.MovePrevious
If RS1.BOF = True Then
MsgBox "NO RECORDS"
RS1.MoveFirst
Call DIS
Text8.SetFocus
Else
Call DIS
Text8.SetFocus
End If
End Sub
Private Sub Command3_Click()
RS1.Fields(0) = Text1.Text
RS1.Fields(1) = Text2.Text
RS1.Fields(2) = Text3.Text
RS1.Fields(3) = Text4.Text
RS1.Fields(4) = Text5.Text
RS1.Fields(5) = Text7.Text
RS1.Update
MsgBox "rec updated"
End Sub
Private Sub Form_Load()
Set WS = DBEngine.Workspaces(0)
Set DB = WS.OpenDatabase("c:\program files\nani\stud.mdb", False, False)
Set RS = DB.OpenRecordset("hostel_info", dbOpenDynaset)
Set Data1.Recordset = RS
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Text7.SetFocus
End If
End Sub
Private Sub text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Text3.SetFocus
End If
End Sub
Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Text2.SetFocus
End If
End Sub
Private Sub text5_KeyPress(KeyAscii As Integer)
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
Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Text5.SetFocus
End If
End Sub

Public Sub MODI()
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM HOSTEL_INFO WHERE HOSTEL_NAME=" & "'" & InputBox(" ENTER HOSTEL NAME") & "'"
Set RS1 = QDF.OpenRecordset()
If RS1.RecordCount = 0 Then
MsgBox "hostel with that name does not exist"
Else
Call DIS
Text8.SetFocus
End If
End Sub

Public Sub DIS()
Text1.Text = RS1(0)
Text2.Text = RS1(1)
Text3.Text = RS1(2)
Text4.Text = RS1(3)
Text5.Text = RS1(4)
Text7.Text = RS1(5)
End Sub
Private Sub TEXT8_KeyPress(KeyAscii As Integer)
If KeyAscii = 68 Or KeyAscii = 100 Then
KeyAscii = 0
If MsgBox("DO U WANT TO DELETE THE REC", vbYesNo) = vbYes Then
RS1.Delete
End If
Else
If KeyAscii = 77 Or KeyAscii = 109 Then
RS1.Edit
Text1.SetFocus
End If
End If
End Sub
Public Sub Save()
RS.Fields(0) = Text1.Text
RS.Fields(1) = Text2.Text
RS.Fields(2) = Text3.Text
RS.Fields(3) = Text4.Text
RS.Fields(4) = Text5.Text
RS.Fields(5) = Val(Text7.Text)
RS.Update
End Sub

Public Sub DEL()
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM HOSTEL_INFO WHERE HOSTEL_NAME=" & "'" & InputBox(" ENTER HOSTEL NAME") & "'"
Set RS1 = QDF.OpenRecordset()
If RS1.RecordCount = 0 Then
MsgBox "hostel with that name does not exist"
Else
Call DIS
Text8.SetFocus
End If
End Sub

Private Sub TEXT7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Text4.SetFocus
End If
End Sub

