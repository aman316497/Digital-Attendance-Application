VERSION 5.00
Begin VB.Form Form18 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digital Attendace Application"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14895
   ControlBox      =   0   'False
   LinkTopic       =   "Form18"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   14895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command63 
      Caption         =   "61"
      Height          =   495
      Left            =   10680
      TabIndex        =   62
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command62 
      Caption         =   "56"
      Height          =   495
      Left            =   9000
      TabIndex        =   61
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command61 
      Caption         =   "57"
      Height          =   495
      Left            =   9000
      TabIndex        =   60
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Command60 
      Caption         =   "58"
      Height          =   495
      Left            =   9000
      TabIndex        =   59
      Top             =   6120
      Width           =   615
   End
   Begin VB.CommandButton Command59 
      Caption         =   "59"
      Height          =   495
      Left            =   9000
      TabIndex        =   58
      Top             =   6840
      Width           =   615
   End
   Begin VB.CommandButton Command58 
      Caption         =   "60"
      Height          =   495
      Left            =   9000
      TabIndex        =   57
      Top             =   7560
      Width           =   615
   End
   Begin VB.CommandButton Command57 
      Caption         =   "51"
      Height          =   495
      Left            =   9000
      TabIndex        =   56
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command56 
      Caption         =   "52"
      Height          =   495
      Left            =   9000
      TabIndex        =   55
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command55 
      Caption         =   "53"
      Height          =   495
      Left            =   9000
      TabIndex        =   54
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command54 
      Caption         =   "54"
      Height          =   495
      Left            =   9000
      TabIndex        =   53
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command53 
      Caption         =   "55"
      Height          =   495
      Left            =   9000
      TabIndex        =   52
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command52 
      Caption         =   "46"
      Height          =   495
      Left            =   7320
      TabIndex        =   51
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command51 
      Caption         =   "47"
      Height          =   495
      Left            =   7320
      TabIndex        =   50
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Command50 
      Caption         =   "48"
      Height          =   495
      Left            =   7320
      TabIndex        =   49
      Top             =   6120
      Width           =   615
   End
   Begin VB.CommandButton Command49 
      Caption         =   "49"
      Height          =   495
      Left            =   7320
      TabIndex        =   48
      Top             =   6840
      Width           =   615
   End
   Begin VB.CommandButton Command48 
      Caption         =   "50"
      Height          =   495
      Left            =   7320
      TabIndex        =   47
      Top             =   7560
      Width           =   615
   End
   Begin VB.CommandButton Command47 
      Caption         =   "41"
      Height          =   495
      Left            =   7320
      TabIndex        =   46
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command46 
      Caption         =   "42"
      Height          =   495
      Left            =   7320
      TabIndex        =   45
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command45 
      Caption         =   "43"
      Height          =   495
      Left            =   7320
      TabIndex        =   44
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command44 
      Caption         =   "44"
      Height          =   495
      Left            =   7320
      TabIndex        =   43
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command43 
      Caption         =   "45"
      Height          =   495
      Left            =   7320
      TabIndex        =   42
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command42 
      Caption         =   "36"
      Height          =   495
      Left            =   5640
      TabIndex        =   41
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command41 
      Caption         =   "37"
      Height          =   495
      Left            =   5640
      TabIndex        =   40
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Command40 
      Caption         =   "38"
      Height          =   495
      Left            =   5640
      TabIndex        =   39
      Top             =   6120
      Width           =   615
   End
   Begin VB.CommandButton Command39 
      Caption         =   "39"
      Height          =   495
      Left            =   5640
      TabIndex        =   38
      Top             =   6840
      Width           =   615
   End
   Begin VB.CommandButton Command38 
      Caption         =   "40"
      Height          =   495
      Left            =   5640
      TabIndex        =   37
      Top             =   7560
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   12000
      Top             =   8880
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12360
      Top             =   8880
   End
   Begin VB.CommandButton Command37 
      Caption         =   "Submit"
      Height          =   495
      Left            =   12960
      TabIndex        =   36
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command36 
      Caption         =   "Reset"
      Height          =   495
      Left            =   12960
      TabIndex        =   35
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command35 
      Caption         =   "35"
      Height          =   495
      Left            =   5640
      TabIndex        =   34
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command34 
      Caption         =   "34"
      Height          =   495
      Left            =   5640
      TabIndex        =   33
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command33 
      Caption         =   "33"
      Height          =   495
      Left            =   5640
      TabIndex        =   32
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command32 
      Caption         =   "32"
      Height          =   495
      Left            =   5640
      TabIndex        =   31
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command31 
      Caption         =   "31"
      Height          =   495
      Left            =   5640
      TabIndex        =   30
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command30 
      Caption         =   "30"
      Height          =   495
      Left            =   3960
      TabIndex        =   29
      Top             =   7560
      Width           =   615
   End
   Begin VB.CommandButton Command29 
      Caption         =   "29"
      Height          =   495
      Left            =   3960
      TabIndex        =   28
      Top             =   6840
      Width           =   615
   End
   Begin VB.CommandButton Command28 
      Caption         =   "28"
      Height          =   495
      Left            =   3960
      TabIndex        =   27
      Top             =   6120
      Width           =   615
   End
   Begin VB.CommandButton Command27 
      Caption         =   "27"
      Height          =   495
      Left            =   3960
      TabIndex        =   26
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Command26 
      Caption         =   "26"
      Height          =   495
      Left            =   3960
      TabIndex        =   25
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command25 
      Caption         =   "25"
      Height          =   495
      Left            =   3960
      TabIndex        =   24
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command24 
      Caption         =   "24"
      Height          =   495
      Left            =   3960
      TabIndex        =   23
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command23 
      Caption         =   "23"
      Height          =   495
      Left            =   3960
      TabIndex        =   22
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command22 
      Caption         =   "22"
      Height          =   495
      Left            =   3960
      TabIndex        =   21
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command21 
      Caption         =   "21"
      Height          =   495
      Left            =   3960
      TabIndex        =   20
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command20 
      Caption         =   "20"
      Height          =   495
      Left            =   2280
      TabIndex        =   19
      Top             =   7560
      Width           =   615
   End
   Begin VB.CommandButton Command19 
      Caption         =   "19"
      Height          =   495
      Left            =   2280
      TabIndex        =   18
      Top             =   6840
      Width           =   615
   End
   Begin VB.CommandButton Command18 
      Caption         =   "18"
      Height          =   495
      Left            =   2280
      TabIndex        =   17
      Top             =   6120
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      Caption         =   "17"
      Height          =   495
      Left            =   2280
      TabIndex        =   16
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      Caption         =   "16"
      Height          =   495
      Left            =   2280
      TabIndex        =   15
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Caption         =   "15"
      Height          =   495
      Left            =   2280
      TabIndex        =   14
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "14"
      Height          =   495
      Left            =   2280
      TabIndex        =   13
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "13"
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "12"
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   "11"
      Height          =   495
      Left            =   2280
      TabIndex        =   10
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "10"
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   7560
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   6840
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   6120
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "Developed by SSAD Group"
      Height          =   495
      Left            =   12720
      TabIndex        =   141
      Top             =   8880
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label79 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   11520
      TabIndex        =   140
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label78 
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
      Left            =   10680
      TabIndex        =   139
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label77 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   9840
      TabIndex        =   138
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label76 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   9840
      TabIndex        =   137
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label75 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   9840
      TabIndex        =   136
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label74 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   9840
      TabIndex        =   135
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label73 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   9840
      TabIndex        =   134
      Top             =   7560
      Width           =   375
   End
   Begin VB.Label Label72 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   9840
      TabIndex        =   133
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label71 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   9840
      TabIndex        =   132
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label70 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   9840
      TabIndex        =   131
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label69 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   9840
      TabIndex        =   130
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label68 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   9840
      TabIndex        =   129
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label67 
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
      Left            =   9000
      TabIndex        =   128
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label66 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   8160
      TabIndex        =   127
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label65 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   8160
      TabIndex        =   126
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label64 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   8160
      TabIndex        =   125
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label63 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   8160
      TabIndex        =   124
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label62 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   8160
      TabIndex        =   123
      Top             =   7560
      Width           =   375
   End
   Begin VB.Label Label61 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   8160
      TabIndex        =   122
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label60 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   8160
      TabIndex        =   121
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label59 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   8160
      TabIndex        =   120
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label58 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   8160
      TabIndex        =   119
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label57 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   8160
      TabIndex        =   118
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label56 
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
      Left            =   7320
      TabIndex        =   117
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label55 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   6480
      TabIndex        =   116
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label54 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   6480
      TabIndex        =   115
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label53 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   6480
      TabIndex        =   114
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label52 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   6480
      TabIndex        =   113
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label51 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   6480
      TabIndex        =   112
      Top             =   7560
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
      Left            =   1080
      TabIndex        =   111
      Top             =   8400
      Width           =   10455
   End
   Begin VB.Label Label49 
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
      Left            =   13440
      TabIndex        =   110
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label Label48 
      Alignment       =   2  'Center
      Caption         =   "00"
      Height          =   375
      Left            =   12840
      TabIndex        =   109
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      Caption         =   "00"
      Height          =   375
      Left            =   13680
      TabIndex        =   108
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      Caption         =   "Sec"
      Height          =   375
      Left            =   13680
      TabIndex        =   107
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      Caption         =   "Min"
      Height          =   375
      Left            =   12840
      TabIndex        =   106
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      Caption         =   "Login Session"
      Height          =   255
      Left            =   12840
      TabIndex        =   105
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      Caption         =   $"Form18.frx":0000
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
      Height          =   3135
      Left            =   12720
      TabIndex        =   104
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label42 
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
      Left            =   12720
      TabIndex        =   103
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label41 
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
      Left            =   12720
      TabIndex        =   102
      Top             =   600
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      Height          =   8055
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   12135
   End
   Begin VB.Label Label40 
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
      TabIndex        =   101
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label39 
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
      TabIndex        =   100
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label38 
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
      TabIndex        =   99
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label37 
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
      TabIndex        =   98
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   6480
      TabIndex        =   97
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   6480
      TabIndex        =   96
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   6480
      TabIndex        =   95
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   6480
      TabIndex        =   94
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   6480
      TabIndex        =   93
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   92
      Top             =   7560
      Width           =   375
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   91
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   90
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   89
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   88
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   87
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   86
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   85
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   84
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4800
      TabIndex        =   83
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   82
      Top             =   7560
      Width           =   375
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   81
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   80
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   79
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   78
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   77
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   76
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   75
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   74
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3120
      TabIndex        =   73
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   72
      Top             =   7560
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   71
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   70
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   69
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   68
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   67
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   66
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   1440
      TabIndex        =   65
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   64
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   63
      Top             =   1080
      Width           =   375
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim all, p, a As Integer
Dim od1, ot1 As Date

Private Sub Command36_Click()
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
Label31.Caption = " "
Label32.Caption = " "
Label33.Caption = " "
Label34.Caption = " "
Label35.Caption = " "

Label55.Caption = " "
Label54.Caption = " "
Label53.Caption = " "
Label52.Caption = " "
Label51.Caption = " "
Label61.Caption = " "
Label60.Caption = " "
Label59.Caption = " "
Label58.Caption = " "
Label57.Caption = " "
Label66.Caption = " "
Label65.Caption = " "
Label64.Caption = " "
Label63.Caption = " "
Label62.Caption = " "
Label72.Caption = " "
Label71.Caption = " "
Label70.Caption = " "
Label69.Caption = " "
Label68.Caption = " "
Label77.Caption = " "
Label76.Caption = " "
Label75.Caption = " "
Label74.Caption = " "
Label73.Caption = " "

Label79.Caption = " "
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
Private Sub Command31_Click()
Label31.Caption = "P"
End Sub
Private Sub Command32_Click()
Label32.Caption = "P"
End Sub
Private Sub Command33_Click()
Label33.Caption = "P"
End Sub
Private Sub Command34_Click()
Label34.Caption = "P"
End Sub
Private Sub Command35_Click()
Label35.Caption = "P"
End Sub


Private Sub Command42_Click()
Label55.Caption = "P"
End Sub
Private Sub Command41_Click()
Label54.Caption = "P"
End Sub
Private Sub Command40_Click()
Label53.Caption = "P"
End Sub
Private Sub Command39_Click()
Label52.Caption = "P"
End Sub
Private Sub Command38_Click()
Label51.Caption = "P"
End Sub
Private Sub Command47_Click()
Label61.Caption = "P"
End Sub
Private Sub Command46_Click()
Label60.Caption = "P"
End Sub
Private Sub Command45_Click()
Label59.Caption = "P"
End Sub
Private Sub Command44_Click()
Label58.Caption = "P"
End Sub
Private Sub Command43_Click()
Label57.Caption = "P"
End Sub
Private Sub Command52_Click()
Label66.Caption = "P"
End Sub
Private Sub Command51_Click()
Label65.Caption = "P"
End Sub
Private Sub Command50_Click()
Label64.Caption = "P"
End Sub
Private Sub Command49_Click()
Label63.Caption = "P"
End Sub
Private Sub Command48_Click()
Label62.Caption = "P"
End Sub
Private Sub Command57_Click()
Label72.Caption = "P"
End Sub
Private Sub Command56_Click()
Label71.Caption = "P"
End Sub
Private Sub Command55_Click()
Label70.Caption = "P"
End Sub
Private Sub Command54_Click()
Label69.Caption = "P"
End Sub
Private Sub Command53_Click()
Label68.Caption = "P"
End Sub
Private Sub Command62_Click()
Label77.Caption = "P"
End Sub
Private Sub Command61_Click()
Label76.Caption = "P"
End Sub
Private Sub Command60_Click()
Label75.Caption = "P"
End Sub
Private Sub Command59_Click()
Label74.Caption = "P"
End Sub
Private Sub Command58_Click()
Label73.Caption = "P"
End Sub
Private Sub Command63_Click()
Label79.Caption = "P"
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
Private Sub Label31_Click()
Label31.Caption = "A"
End Sub
Private Sub Label32_Click()
Label32.Caption = "A"
End Sub
Private Sub Label33_Click()
Label33.Caption = "A"
End Sub
Private Sub Label34_Click()
Label34.Caption = "A"
End Sub
Private Sub Label35_Click()
Label35.Caption = "A"
End Sub



Private Sub Label55_Click()
Label55.Caption = "A"
End Sub
Private Sub Label54_Click()
Label54.Caption = "A"
End Sub
Private Sub Label53_Click()
Label53.Caption = "A"
End Sub
Private Sub Label52_Click()
Label52.Caption = "A"
End Sub
Private Sub Label51_Click()
Label51.Caption = "A"
End Sub
Private Sub Label61_Click()
Label61.Caption = "A"
End Sub
Private Sub Label60_Click()
Label60.Caption = "A"
End Sub
Private Sub Label59_Click()
Label59.Caption = "A"
End Sub
Private Sub Label58_Click()
Label58.Caption = "A"
End Sub
Private Sub Label57_Click()
Label57.Caption = "A"
End Sub
Private Sub Label66_Click()
Label66.Caption = "A"
End Sub
Private Sub Label65_Click()
Label65.Caption = "A"
End Sub
Private Sub Label64_Click()
Label64.Caption = "A"
End Sub
Private Sub Label63_Click()
Label63.Caption = "A"
End Sub
Private Sub Label62_Click()
Label62.Caption = "A"
End Sub
Private Sub Label72_Click()
Label72.Caption = "A"
End Sub
Private Sub Label71_Click()
Label71.Caption = "A"
End Sub
Private Sub Label70_Click()
Label70.Caption = "A"
End Sub
Private Sub Label69_Click()
Label69.Caption = "A"
End Sub
Private Sub Label68_Click()
Label68.Caption = "A"
End Sub
Private Sub Label77_Click()
Label77.Caption = "A"
End Sub
Private Sub Label76_Click()
Label76.Caption = "A"
End Sub
Private Sub Label75_Click()
Label75.Caption = "A"
End Sub
Private Sub Label74_Click()
Label74.Caption = "A"
End Sub
Private Sub Label73_Click()
Label73.Caption = "A"
End Sub
Private Sub Label79_Click()
Label79.Caption = "A"
End Sub




Private Sub Timer1_Timer()
Label47.Caption = Val(Label47.Caption) + 1
If Val(Label47.Caption) = 60 Then
Label47.Caption = 0
End If
End Sub

Private Sub Timer2_Timer()
Label48.Caption = Val(Label48.Caption) + 1
End Sub

Private Sub Command37_Click()

Label47.Caption = 0
Label48.Caption = 0


If Form11.ch = 2 Then

Select Case Form4.a

Case 0


Form11.Data7.Recordset.MoveFirst


p = 0
a = 0
If Label1.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label1.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         1-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label2.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label2.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         2-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label3.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label3.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         3-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label4.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label4.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         4-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label5.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label5.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         5-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label6.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label6.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         6-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label7.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label7.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         7-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label8.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label8.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         8-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label9.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label9.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         9-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label10.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label10.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         10-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label11.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label11.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         11-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label12.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label12.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         12-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label13.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label13.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         13-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label14.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label14.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         14-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label15.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label15.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         15-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label16.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label16.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         16-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label17.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label17.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         17-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label18.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label18.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         18-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label19.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label19.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         19-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label20.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label20.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         20-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label21.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label21.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         21-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label22.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label22.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         22-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label23.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label23.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         23-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label24.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label24.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         24-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label25.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label25.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         25-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label26.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label26.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         26-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label27.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label27.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         27-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label28.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label28.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         28-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label29.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label29.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         29-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label30.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label30.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         30-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label31.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label31.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         31-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label32.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label32.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         32-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label33.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label33.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         33-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label34.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label34.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         34-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label35.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label35.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         35-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If

Form11.Data7.Recordset.MoveNext
If Label55.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label55.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         36-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label54.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label54.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         37-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label53.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label53.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         38-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label52.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label52.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         39-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label51.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label51.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         40-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label61.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label61.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         41-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label60.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label60.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         42-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label59.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label59.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         43-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label58.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label58.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         44-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label57.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label57.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         45-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label66.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label66.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         46-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label65.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label65.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         47-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label64.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label64.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         48-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label63.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label63.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         49-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label62.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label62.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         50-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label72.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label72.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         51-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label71.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label71.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         52-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label70.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label70.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         53-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label69.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label69.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         54-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label68.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label68.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         55-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label77.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label77.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         56-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label76.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label76.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         57-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label75.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label75.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         58-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label74.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label74.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         59-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label73.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label73.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         60-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update
End If
Form11.Data7.Recordset.MoveNext
If Label79.Caption = "P" Then
p = p + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("present").Value = Form11.Data7.Recordset.Fields("present").Value + 1
Form11.Data7.Recordset.Update
ElseIf Label79.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         61-" + Form11.Data7.Recordset.Fields!stu_name
a = a + 1
Form11.Data7.Recordset.Edit
Form11.Data7.Recordset.Fields("absent").Value = Form11.Data7.Recordset.Fields("absent").Value + 1
Form11.Data7.Recordset.Update

End If




Case 1




Form11.Data8.Recordset.MoveFirst


p = 0
a = 0
If Label1.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label1.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         1-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label2.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label2.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         2-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label3.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label3.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         3-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label4.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label4.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         4-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label5.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label5.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         5-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label6.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label6.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         6-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label7.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label7.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         7-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label8.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label8.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         8-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label9.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label9.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         9-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label10.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label10.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         10-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label11.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label11.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         11-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label12.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label12.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         12-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label13.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label13.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         13-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label14.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label14.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         14-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label15.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label15.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         15-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label16.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label16.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         16-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label17.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label17.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         17-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label18.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label18.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         18-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label19.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label19.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         19-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label20.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label20.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         20-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label21.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label21.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         21-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label22.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label22.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         22-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label23.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label23.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         23-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label24.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label24.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         24-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label25.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label25.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         25-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label26.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label26.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         26-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label27.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label27.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         27-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label28.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label28.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         28-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label29.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label29.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         29-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label30.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label30.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         30-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label31.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label31.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         31-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label32.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label32.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         32-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label33.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label33.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         33-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label34.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label34.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         34-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label35.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label35.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         35-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If

Form11.Data8.Recordset.MoveNext
If Label55.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label55.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         36-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label54.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label54.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         37-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label53.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label53.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         38-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label52.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label52.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         39-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label51.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label51.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         40-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label61.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label61.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         41-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label60.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label60.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         42-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label59.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label59.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         43-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label58.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label58.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         44-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label57.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label57.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         45-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label66.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label66.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         46-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label65.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label65.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         47-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label64.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label64.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         48-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label63.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label63.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         49-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label62.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label62.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         50-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label72.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label72.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         51-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label71.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label71.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         52-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label70.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label70.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         53-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label69.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label69.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         54-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label68.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label68.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         55-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label77.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label77.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         56-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label76.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label76.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         57-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label75.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label75.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         58-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label74.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label74.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         59-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label73.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label73.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         60-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update
End If
Form11.Data8.Recordset.MoveNext
If Label79.Caption = "P" Then
p = p + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("present").Value = Form11.Data8.Recordset.Fields("present").Value + 1
Form11.Data8.Recordset.Update
ElseIf Label79.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         61-" + Form11.Data8.Recordset.Fields!stu_name
a = a + 1
Form11.Data8.Recordset.Edit
Form11.Data8.Recordset.Fields("absent").Value = Form11.Data8.Recordset.Fields("absent").Value + 1
Form11.Data8.Recordset.Update

End If


Case 2


Form11.Data9.Recordset.MoveFirst


p = 0
a = 0
If Label1.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label1.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         1-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label2.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label2.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         2-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label3.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label3.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         3-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label4.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label4.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         4-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label5.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label5.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         5-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label6.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label6.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         6-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label7.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label7.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         7-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label8.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label8.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         8-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label9.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label9.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         9-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label10.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label10.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         10-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label11.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label11.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         11-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label12.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label12.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         12-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label13.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label13.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         13-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label14.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label14.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         14-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label15.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label15.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         15-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label16.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label16.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         16-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label17.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label17.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         17-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label18.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label18.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         18-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label19.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label19.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         19-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label20.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label20.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         20-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label21.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label21.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         21-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label22.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label22.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         22-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label23.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label23.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         23-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label24.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label24.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         24-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label25.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label25.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         25-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label26.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label26.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         26-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label27.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label27.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         27-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label28.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label28.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         28-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label29.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label29.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         29-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label30.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label30.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         30-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label31.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label31.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         31-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label32.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label32.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         32-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label33.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label33.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         33-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label34.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label34.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         34-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label35.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label35.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         35-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If

Form11.Data9.Recordset.MoveNext
If Label55.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label55.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         36-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label54.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label54.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         37-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label53.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label53.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         38-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label52.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label52.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         39-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label51.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label51.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         40-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label61.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label61.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         41-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label60.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label60.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         42-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label59.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label59.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         43-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label58.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label58.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         44-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label57.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label57.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         45-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label66.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label66.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         46-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label65.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label65.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         47-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label64.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label64.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         48-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label63.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label63.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         49-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label62.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label62.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         50-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label72.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label72.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         51-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label71.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label71.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         52-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label70.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label70.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         53-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label69.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label69.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         54-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label68.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label68.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         55-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label77.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label77.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         56-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label76.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label76.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         57-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label75.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label75.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         58-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label74.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label74.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         59-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label73.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label73.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         60-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update
End If
Form11.Data9.Recordset.MoveNext
If Label79.Caption = "P" Then
p = p + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("present").Value = Form11.Data9.Recordset.Fields("present").Value + 1
Form11.Data9.Recordset.Update
ElseIf Label79.Caption = "A" Then
Form9.Text1.Text = Form9.Text1.Text + "         61-" + Form11.Data9.Recordset.Fields!stu_name
a = a + 1
Form11.Data9.Recordset.Edit
Form11.Data9.Recordset.Fields("absent").Value = Form11.Data9.Recordset.Fields("absent").Value + 1
Form11.Data9.Recordset.Update

End If




End Select


End If





Timer1.Enabled = False
Timer2.Enabled = False

od1 = DateValue(Now)
ot1 = TimeValue(Now)
Form8.Text9.Text = od1
Form8.Text14.Text = ot1




Form8.Text1.Text = a
Form8.Text2.Text = p

all = a + p
If all < 61 Then
m = MsgBox("Please Complete the all Attendace of the Student", vbOKOnly, "Warning")
Else
Me.Hide
Form8.Show
End If
End Sub







