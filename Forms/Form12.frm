VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digital Attendance Application"
   ClientHeight    =   10635
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   20160
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "Form12.frx":0000
   ScaleHeight     =   10635
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   2295
      Left            =   4080
      TabIndex        =   21
      Top             =   360
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton Command12 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3840
         TabIndex        =   23
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Close |X|"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Close |X|"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   18000
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton Command10 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   18
         Top             =   6360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command9 
         Caption         =   "About Option"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   17
         Top             =   5520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command8 
         Caption         =   "SignUp Option"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   16
         Top             =   4680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Search Option"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   15
         Top             =   3840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Register Option"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   14
         Top             =   3000
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Open Database Option"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   13
         Top             =   2160
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Restart Option"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Start "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   16200
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sign Up"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   13680
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton Command2 
         Caption         =   "Create"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   9
         Top             =   3600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form12.frx":1C121
         Left            =   2520
         List            =   "Form12.frx":1C123
         TabIndex        =   7
         Text            =   "----------------------------------"
         Top             =   2760
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Database\Sign_Up.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "account"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Choose Image"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   2760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "UserID"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   3
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4440
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000002&
      Caption         =   "Developed by SSAD Group  "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   1
      Top             =   10440
      Width           =   20415
   End
   Begin VB.Menu Restart 
      Caption         =   "Restart"
   End
   Begin VB.Menu Open 
      Caption         =   "Open Database"
   End
   Begin VB.Menu Register 
      Caption         =   "Register"
      Begin VB.Menu Science 
         Caption         =   "Science"
         Begin VB.Menu BCA 
            Caption         =   "BCA"
         End
         Begin VB.Menu BSC 
            Caption         =   "BSC"
         End
      End
      Begin VB.Menu Commerce 
         Caption         =   "Commerce"
         Begin VB.Menu BCom 
            Caption         =   "BCom"
         End
      End
      Begin VB.Menu Arts 
         Caption         =   "Arts"
         Begin VB.Menu BA 
            Caption         =   "BA"
         End
      End
   End
   Begin VB.Menu Search 
      Caption         =   "Search"
   End
   Begin VB.Menu SignUp 
      Caption         =   "Sign Up"
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
   Begin VB.Menu Exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sel As Integer
Private Sub About_Click()
Form15.Show

End Sub


Private Sub BA_Click()
sel = 3
Form16.Show
End Sub

Private Sub BCA_Click()
sel = 0
Form16.Show
End Sub

Private Sub BCom_Click()
sel = 2
Form16.Show
End Sub

Private Sub BSC_Click()
sel = 1
Form16.Show
End Sub

Private Sub Command1_Click()
Form1.Show
Form1.Text1.Text = ""
Form2.Text2.Text = ""
Form2.Text1.SelText = ""
Form11.Command1.Visible = True
Form11.Command2.Visible = True
Form11.Command3.Visible = True
Form11.Command4.Visible = False
Form11.Command5.Visible = False
Form11.Command6.Visible = False
Form11.Command7.Visible = False
Form11.Command8.Visible = False
Form4.Command1.Visible = True
Form4.Command2.Visible = True
Form4.Command3.Visible = True
Form4.Command4.Visible = False
Form4.Label3.Visible = False
Form4.Label2.Visible = False
Form4.Text1.Visible = False
Form4.Text1.Text = ""
Form5.Label1.Caption = " "
Form5.Label2.Caption = " "
Form5.Label3.Caption = " "
Form5.Label4.Caption = " "
Form5.Label5.Caption = " "
Form5.Label6.Caption = " "
Form5.Label7.Caption = " "
Form5.Label8.Caption = " "
Form5.Label9.Caption = " "
Form5.Label10.Caption = " "
Form5.Label11.Caption = " "
Form5.Label12.Caption = " "
Form5.Label13.Caption = " "
Form5.Label14.Caption = " "
Form5.Label15.Caption = " "
Form5.Label16.Caption = " "
Form5.Label17.Caption = " "
Form5.Label18.Caption = " "
Form5.Label19.Caption = " "
Form5.Label20.Caption = " "
Form5.Label21.Caption = " "
Form5.Label22.Caption = " "
Form5.Label23.Caption = " "
Form5.Label24.Caption = " "
Form5.Label25.Caption = " "
Form5.Label26.Caption = " "
Form5.Label27.Caption = " "
Form5.Label28.Caption = " "
Form5.Label29.Caption = " "
Form5.Label30.Caption = " "
Form5.Label31.Caption = " "
Form5.Label32.Caption = " "
Form5.Label33.Caption = " "
Form5.Label34.Caption = " "
Form5.Label35.Caption = " "

Form5.Label55.Caption = " "
Form5.Label54.Caption = " "
Form5.Label53.Caption = " "
Form5.Label52.Caption = " "
Form5.Label51.Caption = " "
Form5.Label61.Caption = " "
Form5.Label60.Caption = " "
Form5.Label59.Caption = " "
Form5.Label58.Caption = " "
Form5.Label57.Caption = " "
Form5.Label66.Caption = " "
Form5.Label65.Caption = " "
Form5.Label64.Caption = " "
Form5.Label63.Caption = " "
Form5.Label62.Caption = " "
Form5.Label72.Caption = " "
Form5.Label71.Caption = " "
Form5.Label70.Caption = " "
Form5.Label69.Caption = " "
Form5.Label68.Caption = " "
Form5.Label77.Caption = " "
Form5.Label76.Caption = " "
Form5.Label75.Caption = " "
Form5.Label74.Caption = " "
Form5.Label73.Caption = " "
Form5.Label79.Caption = " "


Form6.Label1.Caption = " "
Form6.Label2.Caption = " "
Form6.Label3.Caption = " "
Form6.Label4.Caption = " "
Form6.Label5.Caption = " "
Form6.Label6.Caption = " "
Form6.Label7.Caption = " "
Form6.Label8.Caption = " "
Form6.Label9.Caption = " "
Form6.Label10.Caption = " "
Form6.Label11.Caption = " "
Form6.Label12.Caption = " "
Form6.Label13.Caption = " "
Form6.Label14.Caption = " "
Form6.Label15.Caption = " "
Form6.Label16.Caption = " "
Form6.Label17.Caption = " "
Form6.Label18.Caption = " "
Form6.Label19.Caption = " "
Form6.Label20.Caption = " "
Form6.Label21.Caption = " "
Form6.Label22.Caption = " "
Form6.Label23.Caption = " "
Form6.Label24.Caption = " "
Form6.Label25.Caption = " "
Form6.Label26.Caption = " "
Form6.Label27.Caption = " "
Form6.Label28.Caption = " "
Form6.Label29.Caption = " "
Form6.Label30.Caption = " "
Form6.Label44.Caption = " "
Form6.Label45.Caption = " "
Form6.Label46.Caption = " "
Form6.Label47.Caption = " "
Form6.Label48.Caption = " "



Form7.Label1.Caption = " "
Form7.Label2.Caption = " "
Form7.Label3.Caption = " "
Form7.Label4.Caption = " "
Form7.Label5.Caption = " "
Form7.Label6.Caption = " "
Form7.Label7.Caption = " "
Form7.Label8.Caption = " "
Form7.Label9.Caption = " "
Form7.Label10.Caption = " "
Form7.Label11.Caption = " "
Form7.Label12.Caption = " "
Form7.Label13.Caption = " "
Form7.Label14.Caption = " "
Form7.Label15.Caption = " "
Form7.Label16.Caption = " "
Form7.Label17.Caption = " "
Form7.Label18.Caption = " "
Form7.Label19.Caption = " "
Form7.Label20.Caption = " "
Form7.Label21.Caption = " "
Form7.Label22.Caption = " "
Form7.Label23.Caption = " "
Form7.Label24.Caption = " "
Form7.Label25.Caption = " "
Form7.Label26.Caption = " "


Form13.Label1.Caption = " "
Form13.Label2.Caption = " "
Form13.Label3.Caption = " "
Form13.Label4.Caption = " "
Form13.Label5.Caption = " "
Form13.Label6.Caption = " "
Form13.Label7.Caption = " "
Form13.Label8.Caption = " "
Form13.Label9.Caption = " "
Form13.Label10.Caption = " "
Form13.Label11.Caption = " "
Form13.Label12.Caption = " "
Form13.Label13.Caption = " "
Form13.Label14.Caption = " "
Form13.Label15.Caption = " "
Form13.Label16.Caption = " "
Form13.Label17.Caption = " "
Form13.Label18.Caption = " "
Form13.Label19.Caption = " "
Form13.Label20.Caption = " "
Form13.Label21.Caption = " "
Form13.Label22.Caption = " "
Form13.Label23.Caption = " "
Form13.Label24.Caption = " "
Form13.Label25.Caption = " "
Form13.Label26.Caption = " "
Form13.Label27.Caption = " "
Form13.Label28.Caption = " "
Form13.Label29.Caption = " "
Form13.Label30.Caption = " "
Form13.Label31.Caption = " "
Form13.Label32.Caption = " "
Form13.Label33.Caption = " "
Form13.Label34.Caption = " "
Form13.Label35.Caption = " "

Form13.Label55.Caption = " "
Form13.Label54.Caption = " "
Form13.Label53.Caption = " "
Form13.Label52.Caption = " "
Form13.Label51.Caption = " "
Form13.Label61.Caption = " "
Form13.Label60.Caption = " "
Form13.Label59.Caption = " "
Form13.Label58.Caption = " "
Form13.Label57.Caption = " "
Form13.Label66.Caption = " "
Form13.Label65.Caption = " "
Form13.Label64.Caption = " "
Form13.Label63.Caption = " "
Form13.Label62.Caption = " "
Form13.Label72.Caption = " "
Form13.Label71.Caption = " "
Form13.Label70.Caption = " "
Form13.Label69.Caption = " "
Form13.Label68.Caption = " "
Form13.Label77.Caption = " "
Form13.Label76.Caption = " "
Form13.Label75.Caption = " "
Form13.Label74.Caption = " "
Form13.Label73.Caption = " "

Form13.Label79.Caption = " "


Form9.Text1.Text = ""


Form10.Text1.Text = ""
Form10.Text2.Text = ""
Form10.Text3.Text = ""
Form10.Text4.Text = ""
Form10.Text5.Text = ""
Form10.Text6.Text = ""
Form10.Text7.Text = ""
Form10.Text8.Text = ""
Form10.Text9.Text = ""
Form10.Text10.Text = ""
Form10.Text11.Text = ""
Form10.Text12.Text = ""

End Sub

Private Sub Command10_Click()
Frame3.Caption = Command10.Caption
Frame3.Visible = True
Command12.Visible = True
Text3.Visible = True
Text3.Text = "End the Application"

End Sub

Private Sub Command11_Click()
Frame2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command9.Visible = False
Command10.Visible = False
Command11.Visible = False


Command12.Visible = False
Frame3.Visible = False
Frame3.Caption = ""
Text3.Text = ""
Text3.Visible = False

End Sub

Private Sub Command12_Click()
Frame3.Visible = False
Text3.Visible = False
Text3.Text = ""
Command12.Visible = False
End Sub

Private Sub Command19_Click()
Frame1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Text1.Visible = False
Text2.Visible = False
Combo1.Visible = False
Command2.Visible = False

Text1.Text = ""
Text2.PasswordChar = ""
Text2.Text = ""
Text2.PasswordChar = "*"
Combo1.ListIndex = 0
Command19.Visible = False
End Sub

Private Sub Command2_Click()
If Text1.SelLength = 0 & Text2.SelLength = 0 Then
MsgBox "Please Dont Leave it Blank"
Text1.Text = ""
Text1.SetFocus
Text2.Text = ""
Else
Data1.Recordset.Fields("user_id").Value = Text1.Text
Data1.Recordset.Fields("password").Value = Text2.Text
Data1.Recordset.Fields("image").Value = Combo1.Text
Data1.Recordset.Update

Frame1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Text1.Visible = False
Text2.Visible = False
Combo1.Visible = False
Command2.Visible = False

End If
End Sub

Private Sub Command3_Click()
Frame3.Caption = Command3.Caption
Frame3.Visible = True
Command12.Visible = True
Text3.Visible = True
Text3.Text = "Start >> UserID Authentication >> Password Security >> Image Analyzer >> Department Selection >> Class(Year) Selection >> Attendance Sheet >> Attendance Report >> Absent Student List >> Storing Attendance Data into DataBase & Exit "


End Sub

Private Sub Command4_Click()
Frame3.Caption = Command4.Caption
Frame3.Visible = True
Command12.Visible = True
Text3.Visible = True
Text3.Text = "Restart the Appication"
End Sub

Private Sub Command5_Click()
Frame3.Caption = Command5.Caption
Frame3.Visible = True
Command12.Visible = True
Text3.Visible = True
Text3.Text = "Display the Overall Attendace Database of all Department, you can sort the Database by the selected Department as well as Class(Year"
End Sub

Private Sub Command6_Click()
Frame3.Caption = Command6.Caption
Frame3.Visible = True
Command12.Visible = True
Text3.Visible = True
Text3.Text = "Registeration of a New student"
End Sub

Private Sub Command7_Click()
Frame3.Caption = Command7.Caption
Frame3.Visible = True
Command12.Visible = True
Text3.Visible = True
Text3.Text = "Searching of a student for lookup the attendace for his/her"
End Sub

Private Sub Command8_Click()
Frame3.Caption = Command8.Caption
Frame3.Visible = True
Command12.Visible = True
Text3.Visible = True
Text3.Text = "Create a New Login Account, but with Admin PassKey you can Enter into the SignUp Process"

End Sub

Private Sub Command9_Click()
Frame3.Caption = Command9.Caption
Frame3.Visible = True
Command12.Visible = True
Text3.Visible = True
Text3.Text = "Display the Application Developer & Designer Information"

End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub Form_Load()
Combo1.AddItem "Image 1"
Combo1.AddItem "Image 2"
Combo1.AddItem "Image 3"
Combo1.AddItem "Image 4"
Combo1.AddItem "Image 5"
Combo1.AddItem "Image 6"
Combo1.AddItem "Image 7"
Combo1.AddItem "Image 8"
Combo1.AddItem "Image 9"
Combo1.AddItem "Image 10"

End Sub

Private Sub Help_Click()
Frame2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Command6.Visible = True
Command6.Visible = True
Command7.Visible = True
Command8.Visible = True
Command9.Visible = True
Command10.Visible = True
Command11.Visible = True

Command12.Visible = False
Frame3.Visible = False
Frame3.Caption = ""
Text3.Text = ""
Text3.Visible = False


End Sub

Private Sub Open_Click()

Form14.Show
End Sub

Private Sub Restart_Click()
Form1.Refresh
Form2.Refresh
Form3.Refresh
Form4.Refresh
Form5.Refresh
Form6.Refresh
Form7.Refresh
Form8.Refresh
Form9.Refresh
Form10.Refresh
Form11.Refresh
Form12.Refresh
Form13.Refresh
Form14.Refresh
Form15.Refresh


Command1.Enabled = True
End Sub

Private Sub Search_Click()
On Error GoTo msg
msg:
MsgBox "Welocme to Search Engine of Student", vbOKOnly, "Message"
Form17.Show
Exit Sub
End Sub

Private Sub SignUp_Click()
Dim ip As String
ip = InputBox("Enter the Secure Pass-Key for register new login account", "Security Warning")
If ip = "" Then
MsgBox "Dont leave it Blank", vbOKOnly, "Error"
Else
If ip = "RAC@Washim" Then
Frame1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Text1.Visible = True
Text2.Visible = True
Combo1.Visible = True
Command2.Visible = True
Command19.Visible = True
Text1.SetFocus
Data1.Recordset.AddNew
Else
MsgBox "Please Enter the correct the Pass-Key", vbOKOnly, "Invalid Pass-Key"
End If
End If
End Sub
