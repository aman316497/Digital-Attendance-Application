VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digital Attendance Application"
   ClientHeight    =   7785
   ClientLeft      =   5655
   ClientTop       =   1725
   ClientWidth     =   10020
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   10020
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Option10 
      BackColor       =   &H80000002&
      Caption         =   "Image 10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   25
      Top             =   3000
      Width           =   1095
   End
   Begin VB.OptionButton Option9 
      BackColor       =   &H80000002&
      Caption         =   "Image 9"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   24
      Top             =   3000
      Width           =   1095
   End
   Begin VB.OptionButton Option8 
      BackColor       =   &H80000002&
      Caption         =   "Image 8"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   23
      Top             =   6120
      Width           =   1095
   End
   Begin VB.OptionButton Option7 
      BackColor       =   &H80000002&
      Caption         =   "Image 7"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   22
      Top             =   3480
      Width           =   1095
   End
   Begin VB.PictureBox Picture10 
      Height          =   1455
      Left            =   2040
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   21
      Top             =   1440
      Width           =   1815
   End
   Begin VB.PictureBox Picture9 
      Height          =   1455
      Left            =   5880
      Picture         =   "Form3.frx":0545
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   20
      Top             =   1440
      Width           =   1815
   End
   Begin VB.PictureBox Picture8 
      Height          =   1455
      Left            =   3960
      Picture         =   "Form3.frx":11E3
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   19
      Top             =   4560
      Width           =   1815
   End
   Begin VB.PictureBox Picture7 
      Height          =   1455
      Left            =   120
      Picture         =   "Form3.frx":1780
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   18
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8160
      TabIndex        =   17
      Text            =   "Attempt:-"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9120
      TabIndex        =   16
      Text            =   "3"
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   2040
      Picture         =   "Form3.frx":37D5
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   15
      Top             =   4080
      Width           =   1815
   End
   Begin VB.PictureBox Picture6 
      Height          =   1455
      Left            =   3960
      Picture         =   "Form3.frx":4735
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   13
      Top             =   1920
      Width           =   1815
   End
   Begin VB.PictureBox Picture5 
      Height          =   1455
      Left            =   120
      Picture         =   "Form3.frx":4EA4
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   12
      Top             =   4560
      Width           =   1815
   End
   Begin VB.PictureBox Picture4 
      Height          =   1455
      Left            =   7800
      Picture         =   "Form3.frx":5C3E
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   11
      Top             =   4560
      Width           =   1815
   End
   Begin VB.PictureBox Picture3 
      Height          =   1455
      Left            =   7800
      Picture         =   "Form3.frx":6051
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   10
      Top             =   1920
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Height          =   1455
      Left            =   5880
      Picture         =   "Form3.frx":664F
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   9
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
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
      Left            =   8160
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H80000002&
      Caption         =   "Image 6"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   3480
      Width           =   1095
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H80000002&
      Caption         =   "Image 5"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   6120
      Width           =   1095
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H80000002&
      Caption         =   "Image 4"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      Top             =   6120
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H80000002&
      Caption         =   "Image 3"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000002&
      Caption         =   "Image 2"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   5640
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000002&
      Caption         =   "Image 1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   5640
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8280
      Top             =   6960
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
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      Caption         =   $"Form3.frx":6AAC
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
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   6960
      Width           =   8055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Developed by SSAD Group"
      Height          =   495
      Left            =   8280
      TabIndex        =   8
      Top             =   7320
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Last Level of Authentication"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public i As Integer
Private Sub Command1_Click()
Data1.RecordSource = "select * from account where user_id ='" + Form2.Text2.Text + "' And password='" + Form2.Text1.Text + "' And image='" + Text1.Text + "' "
Data1.Refresh
If i = 3 Then

MsgBox "Please Restart the application, You have Many Time Wrong Attempts", vbCritical, "Warning"
Form12.Show
Form12.Command1.Enabled = False
Else
If Text1.Text = "" Then
MsgBox "Dont let this Box Empty", vbInformation, "Error"
Else
If Data1.Recordset.EOF Then
MsgBox "You have Choose a Wrong Image", vbCritical, "Access Denied"


i = i + 1
Text6.Text = Val(Text6.Text) - 1

Else


Form11.Show
Me.Hide
Text1.Text = ""

End If
End If
End If


End Sub



Private Sub Form_Load()
i = 1
End Sub


Private Sub Option1_Click()
Text1.Text = Option1.Caption
End Sub
Private Sub Option2_Click()
Text1.Text = Option2.Caption
End Sub
Private Sub Option3_Click()
Text1.Text = Option3.Caption
End Sub
Private Sub Option4_Click()
Text1.Text = Option4.Caption
End Sub
Private Sub Option5_Click()
Text1.Text = Option5.Caption
End Sub
Private Sub Option6_Click()
Text1.Text = Option6.Caption
End Sub
Private Sub Option7_Click()
Text1.Text = Option7.Caption
End Sub
Private Sub Option8_Click()
Text1.Text = Option8.Caption
End Sub
Private Sub Option9_Click()
Text1.Text = Option9.Caption
End Sub
Private Sub Option10_Click()
Text1.Text = Option10.Caption
End Sub


