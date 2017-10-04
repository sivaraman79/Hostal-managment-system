VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox CrystalReport1 
      Height          =   480
      Left            =   360
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   2160
      Width           =   1200
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WS As Workspace
Dim RS As Recordset
Dim DB As Database
Dim RS1 As Recordset
Dim RS2 As Recordset
Dim QDF As QueryDef

Private Sub list1_Click()
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM ROOM_ALLOC WHERE HOSTELNAME = " & "'" & List1.List(List1.ListIndex) & "'"
Set RS1 = QDF.OpenRecordset()
Set RS2 = DB.OpenRecordset("TEMP", dbOpenDynaset)
Do While RS2.EOF = False
RS2.Delete
RS2.MoveNext
Loop
Do While RS1.EOF = False
RS2.AddNew
RS2(0) = RS1(9)
RS2(1) = RS1(1)
RS2(2) = RS1(2)
RS2(3) = RS1(3)
RS2(4) = RS1(4)
RS2(5) = RS1(5)
RS2(6) = RS1(6)
RS2(7) = RS1(7)
RS2.Update
RS1.MoveNext
Loop

.CrystalReport1.Action = 1
'CrystalReport1.ReportLatestPage
End Sub

Private Sub Form_Load()
Dim m, i As Integer
Set WS = DBEngine.Workspaces(0)
Set DB = WS.OpenDatabase("c:\program files\nani\stud.mdb", False, False)
Set RS = DB.OpenRecordset("HOSTEL_info", dbOpenDynaset)
Do While RS.EOF = False
If List1.ListCount = 0 Then
List1.AddItem RS(0)
Else
K = List1.ListCount
Do While K <> i
If RS(0) = List1.List(i) Then
m = m + 1
End If
i = i + 1
Loop
If m = 0 Then
List1.AddItem RS(0)
End If
End If
i = 0
m = 0
RS.MoveNext
Loop
 End Sub
