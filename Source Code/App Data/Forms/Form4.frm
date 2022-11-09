VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digital Attendance Application"
   ClientHeight    =   6465
   ClientLeft      =   7215
   ClientTop       =   2160
   ClientWidth     =   7290
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7290
   Begin VB.CommandButton Command6 
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
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "OK"
      Height          =   735
      Left            =   4200
      TabIndex        =   6
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   6000
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "IIIrd Year"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   4920
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "IInd Year"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   2880
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ist Year"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   840
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "<<Back"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Enter Your Name"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   5640
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Developed by SSAD Group"
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   6000
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Choose one of the Class "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   480
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   4335
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   6375
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public a, choice As Integer
Public n1, rs, l2 As String
Dim id, it As Date
Private Sub Command1_Click()
Command6.Visible = False
Label3.Visible = True
Command2.Visible = False
Command3.Visible = False
Text1.Visible = True
Label2.Visible = True
Command4.Visible = True
Form8.Text17.Text = Command1.Caption
a = 0
rs = Form11.sr + "-" + Command1.Caption
l2 = Form11.l1 + "Ist Year\"
End Sub
Private Sub Command2_Click()
Command6.Visible = False
Label3.Visible = True
Command1.Visible = False
Command3.Visible = False
Text1.Visible = True
Label2.Visible = True
Command4.Visible = True
Form8.Text17.Text = Command2.Caption
a = 1
rs = Form11.sr + "-" + Command2.Caption
l2 = Form11.l1 + "IInd Year\"
End Sub
Private Sub Command3_Click()
Command6.Visible = False
Label3.Visible = True
Command2.Visible = False
Command1.Visible = False
Text1.Visible = True
Label2.Visible = True
Command4.Visible = True
Form8.Text17.Text = Command3.Caption
a = 2
rs = Form11.sr + "-" + Command3.Caption
l2 = Form11.l1 + "IIIrd Year\"
End Sub

Private Sub Command4_Click()
If Text1.Text = "" Then
m = MsgBox("You cannot leave this box blank", vbOKOnly, "Warning")
Else
id = DateValue(Now)
it = TimeValue(Now)
Form8.Text6.Text = id
Form8.Text12.Text = it
n1 = Text1.Text
Form8.Text15.Text = n1

choice = Form11.ch
Select Case choice
Case 0
    Form13.Show
    Me.Hide
Case 1
    If a = 0 Then
    Form5.Show
    Me.Hide
    Form5.Timer1.Enabled = True
    Form5.Timer2.Enabled = True
    ElseIf a = 1 Then
    Form6.Show
    Me.Hide
    Form6.Timer1.Enabled = True
    Form6.Timer2.Enabled = True
    ElseIf a = 2 Then
    Form7.Show
    Me.Hide
    Form7.Timer1.Enabled = True
    Form7.Timer2.Enabled = True
    End If
Case 2
    Form18.Show
    Me.Hide
Case 3
    Form19.Show
    Me.Hide
End Select
End If

End Sub


Private Sub Command6_Click()
Form11.Command8.Visible = False
Form11.Command4.Visible = False
Form11.Command5.Visible = False
Form11.Command6.Visible = False
Form11.Command1.Visible = True
Form11.Command2.Visible = True
Form11.Command3.Visible = True
Label2.Visible = False
Command4.Visible = False
Text1.Visible = False
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command6.Visible = False
Text1.Text = ""
Me.Hide
Form11.Show
End Sub

Private Sub Form_Load()
Me.Refresh
End Sub

Private Sub Label3_Click()
Label2.Visible = False
Command4.Visible = False
Text1.Visible = False
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command6.Visible = True
Label3.Visible = False
Text1.Text = ""

End Sub
