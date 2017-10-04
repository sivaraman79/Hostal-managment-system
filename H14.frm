VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   600
      TabIndex        =   15
      Text            =   " "
      Top             =   6120
      Width           =   255
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   441
      TabCaption(0)   =   "ADD OR DEL OR MOD"
      TabPicture(0)   =   "H14.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Line4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line5"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text5"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text4"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Combo1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Command1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Command2"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Command3"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "VIEW"
      TabPicture(1)   =   "H14.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGrid1"
      Tab(1).Control(1)=   "Data1"
      Tab(1).ControlCount=   2
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "H14.frx":0038
         Height          =   3735
         Left            =   -73800
         OleObjectBlob   =   "H14.frx":004C
         TabIndex        =   18
         Top             =   480
         Width           =   9375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "UPDATE"
         Height          =   495
         Left            =   4080
         TabIndex        =   16
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "PREVIOUS"
         Height          =   495
         Left            =   2640
         TabIndex        =   14
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "NEXT"
         Height          =   495
         Left            =   1200
         TabIndex        =   13
         Top             =   3360
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "H14.frx":0A1F
         Left            =   4680
         List            =   "H14.frx":0A2C
         TabIndex        =   12
         Top             =   840
         Width           =   855
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
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Text            =   "   "
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Text            =   " "
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Text            =   " "
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         Text            =   " "
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Text            =   " "
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Line Line5 
         X1              =   480
         X2              =   9480
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line4 
         X1              =   9480
         X2              =   9480
         Y1              =   4320
         Y2              =   5400
      End
      Begin VB.Line Line3 
         X1              =   480
         X2              =   9480
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line2 
         X1              =   480
         X2              =   480
         Y1              =   4320
         Y2              =   5400
      End
      Begin VB.Label Label7 
         Caption         =   "FOR MODIFICATION SELECT THE REC AND PRESS M,AFTER MODIFICATION PRESS UPDATE"
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   4920
         Width           =   7335
      End
      Begin VB.Label Label6 
         Caption         =   " FOR DELETION SELECT THE RECORD AND PRESS D"
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   4440
         Width           =   4335
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   10560
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label1 
         Caption         =   "hostelname"
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "blocktype"
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "blockdesc"
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "roomstartno"
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "roomendno"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   2520
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form3"
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
Text5.SetFocus
End Sub
Public Sub CLR()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
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
Text6.SetFocus
Else
Call DIS
Text6.SetFocus
End If
End Sub
Private Sub Command2_Click()
RS1.MovePrevious
If RS1.BOF = True Then
MsgBox "NO RECORDS"
RS1.MoveFirst
Call DIS
Text6.SetFocus
Else
Call DIS
Text6.SetFocus
End If
End Sub
Private Sub Command3_Click()
RS1.Fields(0) = Text1.Text
RS1.Fields(1) = Text2.Text
RS1.Fields(2) = Text3.Text
RS1.Fields(3) = Text4.Text
RS1.Fields(4) = Text5.Text
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
Text4.SetFocus
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
Text6.SetFocus
End If
End Sub

Public Sub DIS()
Text1.Text = RS1(0)
Text2.Text = RS1(1)
Text3.Text = RS1(2)
Text4.Text = RS1(3)
Text5.Text = RS1(4)
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
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
Text6.SetFocus
End If
End Sub
