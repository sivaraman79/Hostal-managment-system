VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option2 
      Caption         =   "FEMALE"
      Height          =   495
      Left            =   4080
      TabIndex        =   38
      Top             =   4320
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "MALE"
      Height          =   495
      Left            =   4080
      TabIndex        =   37
      Top             =   3960
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   13361
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "STUDENT INFORMATION"
      TabPicture(0)   =   "H13.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Combo1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "MISCELLENOUS INFORMATION"
      TabPicture(1)   =   "H13.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "HOSTELINFORMATION"
      TabPicture(2)   =   "H13.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   3375
         Left            =   -74040
         TabIndex        =   26
         Top             =   1440
         Width           =   7815
         Begin VB.TextBox Text16 
            Height          =   405
            Left            =   5640
            TabIndex        =   36
            Text            =   " "
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox Text15 
            Height          =   405
            Left            =   5640
            TabIndex        =   35
            Text            =   " "
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox Text14 
            Height          =   405
            Left            =   5640
            TabIndex        =   34
            Text            =   " "
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox Text13 
            Height          =   405
            Left            =   1560
            TabIndex        =   33
            Text            =   " "
            Top             =   1440
            Width           =   1815
         End
         Begin VB.TextBox Text12 
            Height          =   375
            Left            =   1560
            TabIndex        =   32
            Text            =   " "
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label16 
            Caption         =   "ROOMNO"
            Height          =   255
            Left            =   4080
            TabIndex        =   31
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label15 
            Caption         =   "BLOCKDESC"
            Height          =   495
            Left            =   4080
            TabIndex        =   30
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label14 
            Caption         =   "BLOCKNAME"
            Height          =   375
            Left            =   4080
            TabIndex        =   29
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "HOSTELNAME"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "HOSTEL_ID"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   1095
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "H13.frx":0054
         Left            =   600
         List            =   "H13.frx":005E
         TabIndex        =   23
         Top             =   480
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   5775
         Left            =   -74640
         TabIndex        =   12
         Top             =   720
         Visible         =   0   'False
         Width           =   9255
         Begin VB.TextBox Text10 
            Height          =   495
            Left            =   2400
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Top             =   2640
            Width           =   3255
         End
         Begin VB.TextBox Text9 
            Height          =   495
            Left            =   2400
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   1920
            Width           =   2895
         End
         Begin VB.TextBox Text8 
            Height          =   375
            Left            =   2400
            TabIndex        =   20
            Text            =   " "
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox Text7 
            Height          =   405
            Left            =   2400
            TabIndex        =   19
            Text            =   " "
            Top             =   3360
            Width           =   2895
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   2400
            TabIndex        =   18
            Text            =   " "
            Top             =   600
            Width           =   5295
         End
         Begin VB.Label Label10 
            Caption         =   "PHONENO"
            Height          =   375
            Left            =   360
            TabIndex        =   17
            Top             =   3480
            Width           =   1935
         End
         Begin VB.Label Label9 
            Caption         =   "BLOODGROUP"
            Height          =   615
            Left            =   240
            TabIndex        =   16
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "LOCAL ADDRESS"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   2760
            Width           =   2295
         End
         Begin VB.Label Label7 
            Caption         =   "PERMANENT ADDRESS"
            Height          =   615
            Left            =   120
            TabIndex        =   14
            Top             =   2040
            Width           =   3735
         End
         Begin VB.Label Label6 
            Caption         =   "GAURDIANNAME"
            Height          =   375
            Left            =   240
            TabIndex        =   13
            Top             =   600
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   4695
         Left            =   600
         TabIndex        =   1
         Top             =   1200
         Visible         =   0   'False
         Width           =   8895
         Begin VB.TextBox Text11 
            Height          =   405
            Left            =   1800
            TabIndex        =   25
            Text            =   " "
            Top             =   3600
            Width           =   1335
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   1800
            TabIndex        =   11
            Text            =   " "
            Top             =   4200
            Width           =   1575
         End
         Begin VB.TextBox Text4 
            Height          =   405
            Left            =   1800
            TabIndex        =   10
            Text            =   " "
            Top             =   3000
            Width           =   1335
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   1800
            TabIndex        =   9
            Text            =   " "
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            Height          =   405
            Left            =   1680
            TabIndex        =   8
            Text            =   " "
            Top             =   1800
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Height          =   405
            Left            =   1680
            TabIndex        =   7
            Text            =   " "
            ToolTipText     =   "enter  reg no"
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label Label11 
            Caption         =   "FATHER NAME"
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   3720
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "MARTIALSTATUS"
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Top             =   4320
            Width           =   2775
         End
         Begin VB.Label Label4 
            Caption         =   "SEX"
            Height          =   495
            Left            =   120
            TabIndex        =   5
            Top             =   3120
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "NAME"
            Height          =   495
            Left            =   120
            TabIndex        =   4
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "REG_NO"
            Height          =   495
            Left            =   120
            TabIndex        =   3
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "APP_NO"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   1920
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "Form2"
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
Dim RS4 As Recordset
Dim RS5 As Recordset
Dim RS6 As Recordset
Dim RS7 As Recordset
Dim QDF As QueryDef
Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
Frame1.Visible = True
Frame2.Visible = True
Call ADD1
End If
End Sub


Private Sub Command2_Click()
Call SAVE1
End Sub
Private Sub Form_Load()
Set WS = DBEngine.Workspaces(0)
Set DB = WS.OpenDatabase("c:\program files\nani\stud.mdb", False, False)
Set RS = DB.OpenRecordset("ITEM_DESC", dbOpenDynaset)
Set RS1 = DB.OpenRecordset("STUD_INFO", dbOpenDynaset)
Set RS2 = DB.OpenRecordset("ITEM_ALLOC", dbOpenDynaset)
Set RS4 = DB.OpenRecordset("HOSTEL_INFO", dbOpenDynaset)
'Do While RS.EOF = False
'List1.AddItem RS(1)
'RS.MoveNext
'Loop
End Sub


Private Sub Option1_Click()
If Option1.Value = True Then
Text4.Text = "MALE"
End If
Text11.SetFocus
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Text4.Text = "FEMALE"
End If
Text11.SetFocus
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
Dim m, K As Integer
Set QDF = DB.CreateQueryDef("")
Set RS5 = DB.OpenRecordset("STUD_INFO", dbOpenDynaset)
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
Else
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM ROOM_ALLOC WHERE REGNO=" & Val(Text1.Text)
Set rs3 = QDF.OpenRecordset()
Set QDF = DB.CreateQueryDef("")
QDF.SQL = "SELECT * FROM STUD_INFO WHERE REG_NO=" & Val(Text1.Text)
Set RS6 = QDF.OpenRecordset()
If RS6.RecordCount = 0 Then
Text2.Text = rs3(0)
Text3.Text = rs3(2)
Text4.Text = rs3(3)
Text12.Text = rs3(9)
Text13.Text = rs3(4)
Text14.Text = rs3(5)
Text15.Text = rs3(6)
Text16.Text = rs3(7)
Option1.SetFocus
Else
MsgBox "ALREADY DETAILS ENTERED"
Text1.Text = ""
End If
End If
Else
If (KeyAscii >= 65 And KeyAscii <= 97) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8) Or (KeyAscii > 90 And KeyAscii < 97) Then
KeyAscii = 0
MsgBox "invalid entry"
End If
End If
End Sub

Public Sub CLR()
Dim K As Integer
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""

End Sub

Public Sub ADD1()
Call CLR
RS1.AddNew
Text1.SetFocus
End Sub

Public Sub SAVE1()
RS1(0) = Val(Text2.Text)
RS1(1) = Val(Text1.Text)
RS1(2) = Text3.Text
RS1(3) = Text4.Text
RS1(4) = Text5.Text
RS1(5) = Text11.Text
RS1(6) = Text6.Text
RS1(7) = Text9.Text
RS1(8) = Text10.Text
RS1(9) = Text8.Text
RS1(10) = Val(Text7.Text)
RS1.Update
MsgBox "SAVED"
End Sub





Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Text7.SetFocus
End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8) Or (KeyAscii > 90 And KeyAscii < 97) Then
KeyAscii = 0
MsgBox "invalid entry"
Else
If (KeyAscii >= 97 And KeyAscii >= 122) Then
KeyAscii = 0
MsgBox "enter only capitals"
Else
If KeyAscii = 13 Then
KeyAscii = 0
Text5.SetFocus
End If

End If
End If
End Sub




Private Sub text5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8) Or (KeyAscii > 90 And KeyAscii < 97) Then
KeyAscii = 0
MsgBox "invalid entry"
Else
If (KeyAscii >= 97 And KeyAscii >= 122) Then
KeyAscii = 0
MsgBox "enter only capitals"
Else
If KeyAscii = 13 Then
KeyAscii = 0
Text6.SetFocus
End If
End If
End If
End Sub



Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8) Or (KeyAscii > 90 And KeyAscii < 97) Then
KeyAscii = 0
MsgBox "invalid entry"
Else
If (KeyAscii >= 97 And KeyAscii >= 122) Then
KeyAscii = 0
MsgBox "enter only capitals"
Else
If KeyAscii = 13 Then
KeyAscii = 0
Text8.SetFocus
End If
End If
End If
End Sub

Private Sub TEXT7_KeyPress(KeyAscii As Integer)
If KeyAscii = 83 Or KeyAscii = 115 Then
KeyAscii = 0
Call SAVE1
Else
If (KeyAscii >= 65 And KeyAscii <= 97) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8) Or (KeyAscii > 90 And KeyAscii < 97) Then
KeyAscii = 0
MsgBox "invalid entry"
End If
End If
End Sub



Private Sub TEXT8_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8) Or (KeyAscii > 90 And KeyAscii < 97) Then
KeyAscii = 0
MsgBox "invalid entry"
Else
If (KeyAscii >= 97 And KeyAscii >= 122) Then
KeyAscii = 0
MsgBox "enter only capitals"
Else
If KeyAscii = 13 Then
KeyAscii = 0
Text9.SetFocus
End If
End If
End If

End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Text10.SetFocus
End If
End Sub
