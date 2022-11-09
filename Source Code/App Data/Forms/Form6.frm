VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digital Attendance Application"
   ClientHeight    =   8925
   ClientLeft      =   5430
   ClientTop       =   1050
   ClientWidth     =   9990
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   9990
   Begin VB.CommandButton Command37 
      Caption         =   "35"
      Height          =   495
      Left            =   5640
      TabIndex        =   79
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton Command36 
      Caption         =   "34"
      Height          =   495
      Left            =   5640
      TabIndex        =   78
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command35 
      Caption         =   "33"
      Height          =   495
      Left            =   5640
      TabIndex        =   77
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command34 
      Caption         =   "32"
      Height          =   495
      Left            =   5640
      TabIndex        =   76
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command33 
      Caption         =   "31"
      Height          =   495
      Left            =   5640
      TabIndex        =   75
      Top             =   840
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6960
      Top             =   8520
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7440
      Top             =   8520
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Submit"
      Height          =   495
      Left            =   8400
      TabIndex        =   61
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Left            =   600
      TabIndex        =   30
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   495
      Left            =   600
      TabIndex        =   29
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   495
      Left            =   600
      TabIndex        =   28
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   495
      Left            =   600
      TabIndex        =   27
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   495
      Left            =   600
      TabIndex        =   26
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   495
      Left            =   600
      TabIndex        =   25
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   495
      Left            =   600
      TabIndex        =   24
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   495
      Left            =   600
      TabIndex        =   23
      Top             =   5880
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   495
      Left            =   600
      TabIndex        =   22
      Top             =   6600
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "10"
      Height          =   495
      Left            =   600
      TabIndex        =   21
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   "11"
      Height          =   495
      Left            =   2280
      TabIndex        =   20
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "12"
      Height          =   495
      Left            =   2280
      TabIndex        =   19
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "13"
      Height          =   495
      Left            =   2280
      TabIndex        =   18
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "14"
      Height          =   495
      Left            =   2280
      TabIndex        =   17
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Caption         =   "15"
      Height          =   495
      Left            =   2280
      TabIndex        =   16
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      Caption         =   "16"
      Height          =   495
      Left            =   2280
      TabIndex        =   15
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      Caption         =   "17"
      Height          =   495
      Left            =   2280
      TabIndex        =   14
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton Command18 
      Caption         =   "18"
      Height          =   495
      Left            =   2280
      TabIndex        =   13
      Top             =   5880
      Width           =   615
   End
   Begin VB.CommandButton Command19 
      Caption         =   "19"
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   6600
      Width           =   615
   End
   Begin VB.CommandButton Command20 
      Caption         =   "20"
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton Command21 
      Caption         =   "21"
      Height          =   495
      Left            =   3960
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command22 
      Caption         =   "22"
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command23 
      Caption         =   "23"
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command24 
      Caption         =   "24"
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command25 
      Caption         =   "25"
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton Command26 
      Caption         =   "26"
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command27 
      Caption         =   "27"
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton Command28 
      Caption         =   "28"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   5880
      Width           =   615
   End
   Begin VB.CommandButton Command29 
      Caption         =   "29"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   6600
      Width           =   615
   End
   Begin VB.CommandButton Command30 
      Caption         =   "30"
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Reset"
      Height          =   495
      Left            =   8400
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Developed by SSAD Group"
      Height          =   495
      Left            =   7800
      TabIndex        =   86
      Top             =   8520
      Width           =   2415
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label49 
      Alignment       =   2  'Center
      Caption         =   "Roll No       A/P"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   85
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label48 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   6480
      TabIndex        =   84
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   6480
      TabIndex        =   83
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   6480
      TabIndex        =   82
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   6480
      TabIndex        =   81
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   6480
      TabIndex        =   80
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label50 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      Caption         =   "Submit this Entry after completing the Lecture"
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
      Height          =   375
      Left            =   960
      TabIndex        =   74
      Top             =   8040
      Width           =   5775
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   73
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      Caption         =   "00"
      Height          =   375
      Left            =   8160
      TabIndex        =   72
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      Caption         =   "00"
      Height          =   375
      Left            =   9000
      TabIndex        =   71
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      Caption         =   "Sec"
      Height          =   375
      Left            =   9000
      TabIndex        =   70
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      Caption         =   "Min"
      Height          =   375
      Left            =   8160
      TabIndex        =   69
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      Caption         =   "Login Session"
      Height          =   255
      Left            =   8160
      TabIndex        =   68
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      Caption         =   "P= Present"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8160
      TabIndex        =   67
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      Caption         =   "A= Absent"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8160
      TabIndex        =   66
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      Caption         =   $"Form6.frx":0000
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
      Height          =   3615
      Left            =   8160
      TabIndex        =   65
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   7815
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   7335
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      Caption         =   "Roll No       A/P"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   64
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      Caption         =   "Roll No       A/P"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   63
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      Caption         =   "Roll No       A/P"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   62
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   60
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   59
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   58
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   57
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   56
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   55
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   54
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   53
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   52
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   51
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   50
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   49
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   48
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   47
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   46
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   45
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   44
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   43
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   42
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   41
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   40
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   39
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   38
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   37
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   36
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   35
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   34
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   33
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   32
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   31
      Top             =   7320
      Width           =   375
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim all, p, a As Integer
Dim od2, ot2 As Date


Private Sub Command31_Click()
Label1.Caption = " "
Label2.Caption = " "
Label3.Caption = " "
Label4.Caption = " "
Label5.Caption = " "
Label6.Caption = " "
Label7.Caption = " "
Label8.Caption = " "
Label9.Caption = " "
Label10.Caption = " "
Label11.Caption = " "
Label12.Caption = " "
Label13.Caption = " "
Label14.Caption = " "
Label15.Caption = " "
Label16.Caption = " "
Label17.Caption = " "
Label18.Caption = " "
Label19.Caption = " "
Label20.Caption = " "
Label21.Caption = " "
Label22.Caption = " "
Label23.Caption = " "
Label24.Caption = " "
Label25.Caption = " "
Label26.Caption = " "
Label27.Caption = " "
Label28.Caption = " "
Label29.Caption = " "
Label30.Caption = " "
Label44.Caption = " "
Label45.Caption = " "
Label46.Caption = " "
Label47.Caption = " "
Label48.Caption = " "
End Sub

Private Sub Command1_Click()
Label1.Caption = "P"
End Sub
Private Sub Command2_Click()
Label2.Caption = "P"
End Sub
Private Sub Command3_Click()
Label3.Caption = "P"
End Sub
Private Sub Command4_Click()
Label4.Caption = "P"
End Sub
Private Sub Command5_Click()
Label5.Caption = "P"
End Sub
Private Sub Command6_Click()
Label6.Caption = "P"
End Sub
Private Sub Command7_Click()
Label7.Caption = "P"
End Sub
Private Sub Command8_Click()
Label8.Caption = "P"
End Sub
Private Sub Command9_Click()
Label9.Caption = "P"
End Sub
Private Sub Command10_Click()
Label10.Caption = "P"
End Sub
Private Sub Command11_Click()
Label11.Caption = "P"
End Sub
Private Sub Command12_Click()
Label12.Caption = "P"
End Sub
Private Sub Command13_Click()
Label13.Caption = "P"
End Sub
Private Sub Command14_Click()
Label14.Caption = "P"
End Sub
Private Sub Command15_Click()
Label15.Caption = "P"
End Sub
Private Sub Command16_Click()
Label16.Caption = "P"
End Sub
Private Sub Command17_Click()
Label17.Caption = "P"
End Sub
Private Sub Command18_Click()
Label18.Caption = "P"
End Sub
Private Sub Command19_Click()
Label19.Caption = "P"
End Sub
Private Sub Command20_Click()
Label20.Caption = "P"
End Sub
Private Sub Command21_Click()
Label21.Caption = "P"
End Sub
Private Sub Command22_Click()
Label22.Caption = "P"
End Sub
Private Sub Command23_Click()
Label23.Caption = "P"
End Sub
Private Sub Command24_Click()
Label24.Caption = "P"
End Sub
Private Sub Command25_Click()
Label25.Caption = "P"
End Sub
Private Sub Command26_Click()
Label26.Caption = "P"
End Sub
Private Sub Command27_Click()
Label27.Caption = "P"
End Sub
Private Sub command28_click()
Label28.Caption = "P"
End Sub
Private Sub Command29_Click()
Label29.Caption = "P"
End Sub
Private Sub Command30_Click()
Label30.Caption = "P"
End Sub
Private Sub Command33_Click()
Label44.Caption = "P"
End Sub
Private Sub Command34_Click()
Label45.Caption = "P"
End Sub
Private Sub Command35_Click()
Label46.Caption = "P"
End Sub
Private Sub Command36_Click()
Label47.Caption = "P"
End Sub
Private Sub Command37_Click()
Label48.Caption = "P"
End Sub





Private Sub Label1_Click()
Label1.Caption = "A"
End Sub
Private Sub Label2_Click()
Label2.Caption = "A"
End Sub
Private Sub Label3_Click()
Label3.Caption = "A"
End Sub
Private Sub Label4_Click()
Label4.Caption = "A"
End Sub
Private Sub Label5_Click()
Label5.Caption = "A"
End Sub
Private Sub Label6_Click()
Label6.Caption = "A"
End Sub
Private Sub Label7_Click()
Label7.Caption = "A"
End Sub
Private Sub Label8_Click()
Label8.Caption = "A"
End Sub
Private Sub Label9_Click()
Label9.Caption = "A"
End Sub
Private Sub Label10_Click()
Label10.Caption = "A"
End Sub
Private Sub Label11_Click()
Label11.Caption = "A"
End Sub
Private Sub Label12_Click()
Label12.Caption = "A"
End Sub
Private Sub Label13_Click()
Label13.Caption = "A"
End Sub
Private Sub Label14_Click()
Label14.Caption = "A"
End Sub
Private Sub Label15_Click()
Label15.Caption = "A"
End Sub
Private Sub Label16_Click()
Label16.Caption = "A"
End Sub
Private Sub Label17_Click()
Label17.Caption = "A"
End Sub
Private Sub Label18_Click()
Label18.Caption = "A"
End Sub
Private Sub Label19_Click()
Label19.Caption = "A"
End Sub
Private Sub Label20_Click()
Label20.Caption = "A"
End Sub
Private Sub Label21_Click()
Label21.Caption = "A"
End Sub
Private Sub Label22_Click()
Label22.Caption = "A"
End Sub
Private Sub Label23_Click()
Label23.Caption = "A"
End Sub
Private Sub Label24_Click()
Label24.Caption = "A"
End Sub
Private Sub Label25_Click()
Label25.Caption = "A"
End Sub
Private Sub Label26_Click()
Label26.Caption = "A"
End Sub
Private Sub Label27_Click()
Label27.Caption = "A"
End Sub
Private Sub Label28_Click()
Label28.Caption = "A"
End Sub
Private Sub Label29_Click()
Label29.Caption = "A"
End Sub
Private Sub Label30_Click()
Label30.Caption = "A"
End Sub
Private Sub Label44_Click()
Label44.Caption = "A"
End Sub
Private Sub Label45_Click()
Label45.Caption = "A"
End Sub
Private Sub Label46_Click()
Label46.Caption = "A"
End Sub
Private Sub Label47_Click()
Label47.Caption = "A"
End Sub
Private Sub Label48_Click()
Label48.Caption = "A"
End Sub



Private Sub Timer1_Timer()
Label41.Caption = Val(Label41.Caption) + 1
If Val(Label41.Caption) = 60 Then
Label41.Caption = 0
End If
End Sub

Private Sub Timer2_Timer()
Label42.Caption = Val(Label42.Caption) + 1
End Sub




Private Sub Command32_Click()

Label41.Caption = 0
Label42.Caption = 0






Form11.Data5.Recordset.MoveFirst




Timer1.Enabled = False
Timer2.Enabled = False


od2 = DateValue(Now)
ot2 = TimeValue(Now)
Form8.Text9.Text = od2
Form8.Text14.Text = ot2


p = 0
a = 0
If Label1.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update

ElseIf Label1.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         1-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label2.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label2.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         2-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label3.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label3.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         3-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label4.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label4.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         4-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label5.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label5.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         5-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label6.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label6.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         6-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label7.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label7.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         7-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label8.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label8.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         8-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label9.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label9.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         9-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label10.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label10.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         10-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label11.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label11.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         11-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label12.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label12.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         12-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label13.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label13.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         13-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label14.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label14.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         14-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label15.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label15.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         15-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label16.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label16.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         16-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label17.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label17.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         17" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label18.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label18.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         18-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label19.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label19.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         19-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label20.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label20.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         20-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label21.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label21.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         21-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label22.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label22.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         22-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label23.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label23.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         23-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label24.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label24.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         24-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label25.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label25.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         25-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label26.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label26.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         26-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label27.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label27.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         27-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label28.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label28.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         28-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label29.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label29.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         29-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label30.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label30.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         30-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label44.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label44.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         31-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label45.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label45.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         32-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label46.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label46.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         33-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label47.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label47.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         34-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form11.Data5.Recordset.MoveNext
If Label48.Caption = "P" Then
p = p + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("present").Value = Form11.Data5.Recordset.Fields("present").Value + 1
Form11.Data5.Recordset.Update
ElseIf Label48.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         35-" + Form11.Data5.Recordset.Fields!stu_name
a = a + 1
Form11.Data5.Recordset.Edit
Form11.Data5.Recordset.Fields("absent").Value = Form11.Data5.Recordset.Fields("absent").Value + 1
Form11.Data5.Recordset.Update
End If
Form8.Text1.Text = a
Form8.Text2.Text = p
all = a + p
If all < 35 Then
m = MsgBox("Please Complete the all Attendace of the Student", vbOKOnly, "Warning")
Else
Me.Hide
Form8.Show
End If
End Sub


