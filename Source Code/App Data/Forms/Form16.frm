VERSION 5.00
Begin VB.Form Form16 
   BackColor       =   &H80000002&
   Caption         =   "Digital Attendance Application"
   ClientHeight    =   7185
   ClientLeft      =   6615
   ClientTop       =   1800
   ClientWidth     =   8655
   ControlBox      =   0   'False
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   8655
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
      Left            =   7200
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   8
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3720
      TabIndex        =   7
      Top             =   3360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3720
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<<Back"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Third Year"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "IInd Year"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ist Year"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2175
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
      Height          =   495
      Left            =   6960
      TabIndex        =   11
      Top             =   6600
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      Caption         =   "Student Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      Caption         =   "Roll No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      Height          =   4335
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Visible         =   0   'False
      Width           =   7695
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sely As Integer
Private Sub Command1_Click()
Shape1.Visible = True
Label1.Visible = True
Label2.Visible = True
Text1.Visible = True
Text2.Visible = True
Command5.Visible = True
Command2.Visible = False
Command3.Visible = False
Command4.Visible = True
sely = 0
Text1.Enabled = False
Text2.Enabled = False
End Sub

Private Sub Command19_Click()

Shape1.Visible = False
Label1.Visible = False
Label2.Visible = False
Text1.Visible = False
Text1.Text = ""
Text2.Visible = False
Text2.Text = ""
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Me.Hide
Form12.Show
End Sub

Private Sub Command2_Click()
Shape1.Visible = True
Label1.Visible = True
Label2.Visible = True
Text1.Visible = True
Text2.Visible = True
Command5.Visible = True
Command1.Visible = False
Command3.Visible = False
Command4.Visible = True
sely = 1

Text1.Enabled = False
Text2.Enabled = False
End Sub

Private Sub Command3_Click()
Shape1.Visible = True
Label1.Visible = True
Label2.Visible = True
Text1.Visible = True
Text2.Visible = True
Command5.Visible = True
Command2.Visible = False
Command1.Visible = False
Command4.Visible = True
sely = 2

Text1.Enabled = False
Text2.Enabled = False
End Sub

Private Sub Command4_Click()

Shape1.Visible = False
Label1.Visible = False
Label2.Visible = False
Text1.Visible = False
Text1.Text = ""
Text2.Visible = False
Text2.Text = ""
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True

End Sub

Private Sub Command5_Click()
On Error Resume Next


If Form12.sel = 0 Then

Select Case sely
Case 0
Form11.Data4.Recordset.MoveLast
Form11.Data4.Recordset.AddNew
Text1.Enabled = True
Text2.Enabled = True
Text1.SetFocus
Case 1
Form11.Data5.Recordset.MoveLast
Form11.Data5.Recordset.AddNew
Text1.Enabled = True
Text2.Enabled = True
Text1.SetFocus
Case 2
Form11.Data6.Recordset.MoveLast
Form11.Data6.Recordset.AddNew
Text1.Enabled = True
Text2.Enabled = True
Text1.SetFocus
End Select


ElseIf Form12.sel = 1 Then

Select Case sely
Case 0
Form11.Data10.Recordset.MoveLast
Form11.Data10.Recordset.AddNew
Text1.Enabled = True
Text2.Enabled = True
Text1.SetFocus
Case 1
Form11.Data11.Recordset.MoveLast
Form11.Data11.Recordset.AddNew
Text1.Enabled = True
Text2.Enabled = True
Text1.SetFocus
Case 2
Form11.Data12.Recordset.MoveLast
Form11.Data12.Recordset.AddNew
Text1.Enabled = True
Text2.Enabled = True
Text1.SetFocus
End Select


ElseIf Form12.sel = 2 Then

Select Case sely
Case 0
Form11.Data7.Recordset.MoveLast
Form11.Data7.Recordset.AddNew
Text1.Enabled = True
Text2.Enabled = True
Text1.SetFocus
Case 1
Form11.Data8.Recordset.MoveLast
Form11.Data8.Recordset.AddNew
Text1.Enabled = True
Text2.Enabled = True
Text1.SetFocus
Case 2
Form11.Data9.Recordset.MoveLast
Form11.Data9.Recordset.AddNew
Text1.Enabled = True
Text2.Enabled = True
Text1.SetFocus
End Select


ElseIf Form12.sel = 3 Then

Select Case sely
Case 0
Form11.Data1.Recordset.MoveLast
Form11.Data1.Recordset.AddNew
Text1.Enabled = True
Text2.Enabled = True
Text1.SetFocus
Case 1
Form11.Data2.Recordset.MoveLast
Form11.Data2.Recordset.AddNew
Text1.Enabled = True
Text2.Enabled = True
Text1.SetFocus
Case 2
Form11.Data3.Recordset.MoveLast
Form11.Data3.Recordset.AddNew
Text1.Enabled = True
Text2.Enabled = True
Text1.SetFocus
End Select

End If




Command6.Visible = True
Command5.Visible = False

End Sub

Private Sub Command6_Click()
On Error Resume Next

If Form12.sel = 0 Then

Select Case sely
Case 0
Form11.Data4.Recordset.Fields("rollno").Value = Text1.Text
Form11.Data4.Recordset.Fields("stu_name").Value = Text2.Text
Form11.Data4.Recordset.Fields("present").Value = 0
Form11.Data4.Recordset.Fields("absent").Value = 0

Form11.Data4.Recordset.Update
Case 1
Form11.Data5.Recordset.Fields("rollno").Value = Text1.Text
Form11.Data5.Recordset.Fields("stu_name").Value = Text2.Text
Form11.Data5.Recordset.Fields("present").Value = 0
Form11.Data5.Recordset.Fields("absent").Value = 0
Form11.Data5.Recordset.Update
Case 2
Form11.Data6.Recordset.Fields("rollno").Value = Text1.Text
Form11.Data6.Recordset.Fields("stu_name").Value = Text2.Text
Form11.Data6.Recordset.Fields("present").Value = 0
Form11.Data6.Recordset.Fields("absent").Value = 0
Form11.Data6.Recordset.Update
End Select


ElseIf Form12.sel = 1 Then

Select Case sely
Case 0
Form11.Data10.Recordset.Fields("rollno").Value = Text1.Text
Form11.Data10.Recordset.Fields("stu_name").Value = Text2.Text
Form11.Data10.Recordset.Fields("present").Value = 0
Form11.Data10.Recordset.Fields("absent").Value = 0
Form11.Data10.Recordset.Update
Case 1
Form11.Data11.Recordset.Fields("rollno").Value = Text1.Text
Form11.Data11.Recordset.Fields("stu_name").Value = Text2.Text
Form11.Data11.Recordset.Fields("present").Value = 0
Form11.Data11.Recordset.Fields("absent").Value = 0
Form11.Data11.Recordset.Update
Case 2
Form11.Data12.Recordset.Fields("rollno").Value = Text1.Text
Form11.Data12.Recordset.Fields("stu_name").Value = Text2.Text
Form11.Data12.Recordset.Fields("present").Value = 0
Form11.Data12.Recordset.Fields("absent").Value = 0
Form11.Data12.Recordset.Update
End Select


ElseIf Form12.sel = 2 Then

Select Case sely
Case 0
Form11.Data7.Recordset.Fields("rollno").Value = Text1.Text
Form11.Data7.Recordset.Fields("stu_name").Value = Text2.Text
Form11.Data7.Recordset.Fields("present").Value = 0
Form11.Data7.Recordset.Fields("absent").Value = 0
Form11.Data7.Recordset.Update
Case 1
Form11.Data8.Recordset.Fields("rollno").Value = Text1.Text
Form11.Data8.Recordset.Fields("stu_name").Value = Text2.Text
Form11.Data8.Recordset.Fields("present").Value = 0
Form11.Data8.Recordset.Fields("absent").Value = 0
Form11.Data8.Recordset.Update
Case 2
Form11.Data9.Recordset.Fields("rollno").Value = Text1.Text
Form11.Data9.Recordset.Fields("stu_name").Value = Text2.Text
Form11.Data9.Recordset.Fields("present").Value = 0
Form11.Data9.Recordset.Fields("absent").Value = 0
Form11.Data9.Recordset.Update
End Select


ElseIf Form12.sel = 3 Then

Select Case sely
Case 0
Form11.Data1.Recordset.Fields("rollno").Value = Text1.Text
Form11.Data1.Recordset.Fields("stu_name").Value = Text2.Text
Form11.Data1.Recordset.Fields("present").Value = 0
Form11.Data1.Recordset.Fields("absent").Value = 0
Form11.Data1.Recordset.Update
Case 1
Form11.Data2.Recordset.Fields("rollno").Value = Text1.Text
Form11.Data2.Recordset.Fields("stu_name").Value = Text2.Text
Form11.Data2.Recordset.Fields("present").Value = 0
Form11.Data2.Recordset.Fields("absent").Value = 0
Form11.Data2.Recordset.Update
Case 2
Form11.Data3.Recordset.Fields("rollno").Value = Text1.Text
Form11.Data3.Recordset.Fields("stu_name").Value = Text2.Text
Form11.Data3.Recordset.Fields("present").Value = 0
Form11.Data3.Recordset.Fields("absent").Value = 0
Form11.Data3.Recordset.Update
End Select

End If




Command7.Visible = True
Command6.Visible = False
MsgBox "New Record has Been Added to Your Database", vbInformation, "Saved"
End Sub


Private Sub Command7_Click()
Text1.Text = ""
Text2.Text = ""

Shape1.Visible = False
Label1.Visible = False
Label2.Visible = False
Text1.Visible = False
Text1.Text = ""
Text2.Visible = False
Text2.Text = ""
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Me.Hide
Form12.Show


End Sub
