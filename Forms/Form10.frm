VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digital Attendance Application"
   ClientHeight    =   9900
   ClientLeft      =   5430
   ClientTop       =   390
   ClientWidth     =   11280
   ControlBox      =   0   'False
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   11280
   Begin VB.CommandButton Command3 
      Caption         =   "Add New Record"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   8760
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Store Record in Database"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   13
      Top             =   8760
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Text12 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   240
      Width           =   8415
   End
   Begin VB.TextBox Text11 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   600
      Width           =   8415
   End
   Begin VB.TextBox Text10 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   8040
      Width           =   8415
   End
   Begin VB.TextBox Text9 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   3840
      Width           =   8415
   End
   Begin VB.TextBox Text8 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   3480
      Width           =   8415
   End
   Begin VB.TextBox Text7 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3120
      Width           =   8415
   End
   Begin VB.TextBox Text6 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2760
      Width           =   8415
   End
   Begin VB.TextBox Text5 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2400
      Width           =   8415
   End
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2040
      Width           =   8415
   End
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   8415
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   8415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   1
      Top             =   8760
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   8415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Developed by SSAD Group"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   19
      Top             =   9360
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      Caption         =   "Step 3:- Finally Click on ""Exit"" to close the window"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   9000
      TabIndex        =   18
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      Caption         =   "Step 2:- Second Click on ""Store Record in Database"" to store the entered into Databse"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1695
      Left            =   9000
      TabIndex        =   17
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      Caption         =   "Step 1:- First Click on ""Add New Record"" to add the data into databse"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1335
      Left            =   9000
      TabIndex        =   16
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      Caption         =   "Note:- 3 Step to Store Data into the Database:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   9000
      TabIndex        =   15
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim msg, sss As String

Private Sub Command1_Click()
On Error Resume Next


msg = MsgBox("Thank You for Using This Digital Attendance Application. Todays Attendance DataSheet has been Save to the Database", vbOKOnly, "Message")
sss = Form4.rs
Open Form4.l2 + sss + ".txt" For Append As #1
Print #1, Text12.Text
Print #1, Text11.Text
Print #1, Text1.Text
Print #1, Text2.Text
Print #1, Text3.Text
Print #1, Text4.Text
Print #1, Text5.Text
Print #1, Text6.Text
Print #1, Text7.Text
Print #1, Text8.Text
Print #1, Text9.Text
Print #1, Text10.Text
Close #1
Me.Hide
Form12.Show

End Sub


Private Sub Command2_Click()

Form14.rec.Fields("DATE").Value = Form8.Text6.Text
Form14.rec.Fields("DEPT").Value = Form8.Text19.Text
Form14.rec.Fields("CLASS").Value = Form8.Text17.Text
Form14.rec.Fields("LECTURAL").Value = Form8.Text15.Text
Form14.rec.Fields("LOGIN_SESSION").Value = Form8.Text12.Text
Form14.rec.Fields("LOGOUT_SESSION").Value = Form8.Text14.Text
Form14.rec.Fields("PRESENT").Value = Form8.Text2.Text
Form14.rec.Fields("ABSENT").Value = Form8.Text1.Text
Form14.rec.Update

Command2.Visible = False
Command1.Visible = True

End Sub

Private Sub Command3_Click()
On Error Resume Next

Form14.rec.AddNew
Command3.Visible = False
Command2.Visible = True


End Sub

Private Sub Form_Load()
On Error Resume Next

Set Form14.DataGrid1.DataSource = Form14.rec


End Sub

Private Sub Text13_Change()

End Sub
