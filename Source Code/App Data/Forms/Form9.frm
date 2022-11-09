VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digital Attendance Application"
   ClientHeight    =   9165
   ClientLeft      =   6315
   ClientTop       =   1500
   ClientWidth     =   8055
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   8055
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   7200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
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
      Height          =   6615
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   7095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Developed by SSAD Group"
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   8760
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      Caption         =   $"Form9.frx":0000
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
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   7920
      Width           =   8055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H80000007&
      Height          =   7695
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim S(12) As String
Private Sub Command1_Click()
Form10.Command3.Visible = True
Form10.Command2.Visible = False
Form10.Command1.Visible = False

S(11) = Form8.de
S(10) = Form8.fd8
S(0) = Form8.fd7
S(1) = Form8.fd6
S(2) = Form8.fd1
S(3) = Form8.fd2
S(4) = Form8.fd3
S(5) = Form8.fd4
S(6) = Form8.fd5
S(7) = "Student Those Who are Absent in the Lecture ::-->"
S(8) = Text1.Text
S(9) = " "
Form10.Text1.Text = S(0)
Form10.Text2.Text = S(1)
Form10.Text3.Text = S(2)
Form10.Text4.Text = S(3)
Form10.Text5.Text = S(4)
Form10.Text6.Text = S(5)
Form10.Text7.Text = S(6)
Form10.Text8.Text = S(7)
Form10.Text9.Text = S(8)
Form10.Text10.Text = S(9)
Form10.Text11.Text = S(10)
Form10.Text12.Text = S(11)

Me.Hide
Form10.Show
End Sub

