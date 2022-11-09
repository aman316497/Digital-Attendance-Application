VERSION 5.00
Begin VB.Form Form15 
   BackColor       =   &H80000002&
   Caption         =   "Digital Attendance Application"
   ClientHeight    =   6555
   ClientLeft      =   4605
   ClientTop       =   1575
   ClientWidth     =   11100
   ControlBox      =   0   'False
   Icon            =   "Form15.frx":0000
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   11100
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   7
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Guidance Teacher"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   2895
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "D.V.Somani Sir"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Designer And Developer  >>>>"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2895
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   360
         Picture         =   "Form15.frx":2A42
         ScaleHeight     =   1785
         ScaleWidth      =   2145
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "SSAD Group"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   2280
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   6015
      Left            =   3120
      Picture         =   "Form15.frx":9257
      ScaleHeight     =   5955
      ScaleWidth      =   7995
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.Line Line7 
         X1              =   5040
         X2              =   5040
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Line Line6 
         X1              =   3720
         X2              =   3720
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Line Line5 
         X1              =   2280
         X2              =   2280
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Line Line4 
         X1              =   5775
         X2              =   5520
         Y1              =   1080
         Y2              =   1575
      End
      Begin VB.Line Line3 
         X1              =   4320
         X2              =   4080
         Y1              =   1080
         Y2              =   1800
      End
      Begin VB.Line Line2 
         X1              =   2880
         X2              =   3135
         Y1              =   1080
         Y2              =   2295
      End
      Begin VB.Line Line1 
         X1              =   1800
         X2              =   2175
         Y1              =   1080
         Y2              =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Deepak              Shubham                   Aman                     Sachin                  "
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   720
         Width           =   5295
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   840
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   6015
      End
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form12.Show
Me.Hide

End Sub
