VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digital Attendance Application"
   ClientHeight    =   7830
   ClientLeft      =   5655
   ClientTop       =   1500
   ClientWidth     =   9150
   ControlBox      =   0   'False
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   9150
   Begin VB.Data Data12 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Database\Stu_Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BSC3rd"
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data11 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Database\Stu_Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BSC2nd"
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data10 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Database\Stu_Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BSC1st"
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data9 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Database\Stu_Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BCOM3rd"
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data8 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Database\Stu_Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BCOM2nd"
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data7 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Database\Stu_Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BCOM1st"
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data6 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Database\Stu_Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BCA3rd"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data5 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Database\Stu_Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BCA2nd"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data4 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Database\Stu_Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BCA1st"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data3 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Database\Stu_Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BA3rd"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Database\Stu_Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BA2nd"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Database\Stu_Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BA1st"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "<< Back"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "B.A."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4320
      TabIndex        =   8
      Top             =   4920
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      Caption         =   "B.COM."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4320
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "B.C.A."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6240
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "B.S.C."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3600
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Commerce"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Science"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Art"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   6135
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   8895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Choose one of the Department"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   7455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Developed by SSAD Group"
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   7320
      Width           =   1695
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ch As Integer
Public ss, sr, l1 As String
Private Sub Command1_Click()
Form8.Text22.Text = Command1.Caption
Command8.Visible = True
Command4.Visible = True
Command5.Visible = True
Command2.Visible = False
Command3.Visible = False
ss = Command1.Caption + "-"
End Sub

Private Sub Command2_Click()
Form8.Text22.Text = Command2.Caption
Command8.Visible = True
Command6.Visible = True
Command1.Visible = False
Command3.Visible = False
ss = Command2.Caption + "-"
End Sub

Private Sub Command3_Click()
Form8.Text22.Text = Command3.Caption
Command8.Visible = True
Command7.Visible = True
Command1.Visible = False
Command2.Visible = False
ss = Command3.Caption + "-"
End Sub

Private Sub Command4_Click()
l1 = "C:\Database\Science\B.S.C\"
Form8.Text19.Text = Command4.Caption
Form4.Show
Me.Hide
Form8.Text17.Text = Command4.Caption + " "
sr = ss + "-" + Command4.Caption
ch = 0
End Sub

Private Sub Command5_Click()
l1 = "C:\Database\Science\B.C.A\"
Form8.Text19.Text = Command5.Caption
Form4.Show
Me.Hide
Form8.Text17.Text = Command5.Caption + " "
sr = ss + "-" + Command5.Caption
ch = 1
End Sub

Private Sub Command6_Click()
l1 = "C:\Database\Commerce\B.Comm\"
Form8.Text19.Text = Command6.Caption
Form4.Show
Me.Hide
Form8.Text17.Text = Command6.Caption + " "
sr = ss + "-" + Command6.Caption
ch = 2
End Sub

Private Sub Command7_Click()
l1 = "C:\Database\Arts\B.A\"
Form8.Text19.Text = Command7.Caption
Form4.Show
Me.Hide
Form8.Text17.Text = Command7.Caption + " "
sr = ss + "-" + Command7.Caption
ch = 3
End Sub

Private Sub Command8_Click()
Command8.Visible = False
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False


End Sub

