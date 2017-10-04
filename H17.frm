VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   13785
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "H17.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(1)=   "DBGrid2"
      Tab(0).Control(2)=   "Data2"
      Tab(0).Control(3)=   "Text1"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGrid1"
      Tab(1).Control(1)=   "Data1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "H17.frx":001C
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "DBGrid3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Text2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Data3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Command1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Text3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Text4"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   960
         TabIndex        =   11
         Text            =   " "
         Top             =   6840
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   960
         TabIndex        =   10
         Text            =   " "
         Top             =   6120
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "DELETE"
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Data Data3 
         Connect         =   "Access"
         DatabaseName    =   "C:\Program Files\nani\Stud.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   7440
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   8400
         TabIndex        =   7
         Text            =   " "
         Top             =   5520
         Width           =   1695
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "H17.frx":0038
         Height          =   2415
         Left            =   240
         OleObjectBlob   =   "H17.frx":004C
         TabIndex        =   6
         Top             =   2400
         Width           =   11415
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   -70680
         TabIndex        =   5
         Text            =   " "
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   "C:\Program Files\nani\Stud.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   615
         Left            =   -72000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6120
         Visible         =   0   'False
         Width           =   3495
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "H17.frx":0A1F
         Height          =   1935
         Left            =   -74880
         OleObjectBlob   =   "H17.frx":0A33
         TabIndex        =   3
         Top             =   3120
         Width           =   11175
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Program Files\nani\Stud.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   -73680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5640
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "H17.frx":1406
         Height          =   2535
         Left            =   -74880
         OleObjectBlob   =   "H17.frx":141A
         TabIndex        =   2
         Top             =   2520
         Width           =   11295
      End
      Begin VB.Label Label3 
         Caption         =   "TOTAL"
         Height          =   375
         Left            =   6360
         TabIndex        =   8
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "ENTER THE REGISTRATION NO"
         Height          =   495
         Left            =   -73800
         TabIndex        =   4
         Top             =   1680
         Width           =   2655
      End
   End
   Begin VB.Label Label1 
      Caption         =   "DELETING RECORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WS As Workspace
Dim DB As Database
Dim RS As Recordset
Dim RS1 As Recordset
Dim RS2 As Recordset
Dim rs3 As Recordset
Dim QDF As QueryDef

Private Sub Command1_Click()
If (MsgBox("HAVE YOU COLLECTED THE ITEMS FROM STUDENT", vbYesNo) = vbYes) Then
Text3.Text = ""
Text3.Text = "y"
If (MsgBox("DID HE PAID THE FEE", vbYesNo) = vbYes) Then
Text4.Text = ""
Text4.Text = "y"
rs3.AddNew
DBGrid2.Row = 0
DBGrid2.Col = 2
rs3.Fields(0) = Val(Text1.Text)
MsgBox DBGrid2.Text
rs3.Fields(1) = DBGrid2.Text
DBGrid2.Col = 4
rs3.Fields(2) = DBGrid2.Text
DBGrid2.Col = 5
rs3.Fields(3) = DBGrid2.Text
DBGrid2.Col = 6
rs3.Fields(4) = DBGrid2.Text
DBGrid2.Col = 7
rs3.Fields(5) = DBGrid2.Text
rs3.Fields(6) = Text3.Text
rs3.Fields(7) = Text4.Text
rs3.Fields(8) = Date
rs3.Update
If (MsgBox("do you want to delete", vbYesNo) = vbYes) Then
If J > 0 Then
RS.Delete
End If
If i > 0 Then
RS1.Delete
End If
If K > 0 Then
RS2.MoveFirst
End If
Do While RS2.EOF = False
RS2.Delete
RS2.MoveNext
Loop
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Else
GoTo 2
End If
Else
MsgBox "collect the fine"
End If
Else
MsgBox "collect the items"
End If
2 End Sub


Private Sub Form_Click()
Text1.SetFocus
End Sub
Private Sub Form_Load()
Set WS = DBEngine.Workspaces(0)
Set DB = WS.OpenDatabase("C:\program files\nani\stud.mdb", False, False)
Set rs3 = DB.OpenRecordset("vaccating_details", dbOpenDynaset)
End Sub



Private Sub text1_KeyPress(KeyAscii As Integer)
Dim F As Integer
Dim i, J, K As Integer
If KeyAscii = 13 Then
KeyAscii = 0
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM  ROOM_ALLOC WHERE REGNO=" & Val(Text1.Text)
Set RS = QDF.OpenRecordset()
If RS.RecordCount = 0 Then
MsgBox "no room allocated on this roomno"
J = J + 1
Else
Set Data2.Recordset = RS
End If
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM  ITEM_ALLOC WHERE REGNO=" & Val(Text1.Text)
Set RS1 = QDF.OpenRecordset()
If RS1.RecordCount = 0 Then
MsgBox "no items alocated to him"
i = i + 1
Else
Set Data1.Recordset = RS1
End If
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM  FINE_DETAILS WHERE REGNO=" & Val(Text1.Text)
Set RS2 = QDF.OpenRecordset()
If RS2.RecordCount = 0 Then
K = K + 1
MsgBox "no fine for him"
Else
Set Data3.Recordset = RS2
End If
Do While RS2.EOF = False
F = F + RS2(7)
RS2.MoveNext
Loop
Text2.Text = Val(F)
End If
End Sub
