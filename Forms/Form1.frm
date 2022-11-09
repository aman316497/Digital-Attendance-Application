VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digital Attendance Application"
   ClientHeight    =   5535
   ClientLeft      =   5880
   ClientTop       =   1935
   ClientWidth     =   9045
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9045
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6360
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Database\Sign_Up.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Database\Sign_Up.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from account"
      Caption         =   "Data1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Text            =   "UserID"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   6000
      Picture         =   "Form1.frx":2A42
      ScaleHeight     =   2355
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Developed by SSAD Group"
      Height          =   495
      Left            =   7440
      TabIndex        =   3
      Top             =   5040
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Rajasthan Aryan Arts,Shri Mithulal Kacholia Commerce and Shri Satyanarayanji Ramkrishnaji Rathi Science College"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   795
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   7485
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   3135
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Text6.Text = 3
Form3.Text6.Text = 3

Data1.RecordSource = "select * from account where user_id='" + Text1.Text + "'"
Data1.Refresh

Form2.Text1.PasswordChar = 0
Form2.Text1.Text = ""
Form2.Text1.PasswordChar = "*"
If Text1.Text = "" Then
MsgBox "Dont let this Box Empty", vbInformation, "Error"
Else

If Data1.Recordset.EOF Then
MsgBox "You Have Enter a Invalid UserId, Please Enter the Correct one", vbCritical, "Message"
Text1.Text = ""
Text1.SetFocus
Else
Form2.Show
Me.Hide
Form2.Text2.Text = Text1.Text
Form2.i = 1
Form3.i = 1
End If
End If
End Sub

Private Sub Form_Load()
Me.Hide
Form12.Show


End Sub


