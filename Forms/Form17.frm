VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form17 
   BackColor       =   &H80000002&
   Caption         =   "Digital Attendance Application"
   ClientHeight    =   5565
   ClientLeft      =   5505
   ClientTop       =   3120
   ClientWidth     =   10935
   ControlBox      =   0   'False
   LinkTopic       =   "Form17"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   10935
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
      Left            =   4920
      TabIndex        =   18
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command17 
      Caption         =   "IIIrd Year"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   17
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      Caption         =   "IInd Year"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Ist Year"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "IIIrd Year"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   14
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "IInd Year"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Ist Year"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   12
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   "IIIrd Year"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "IInd Year"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Ist Year"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "IIIrd Year"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "IInd Year"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Ist Year"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "B.A."
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "B.Comm."
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "B.S.C."
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "B.C.A."
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select Department"
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
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form17.frx":0000
      Height          =   4815
      Left            =   6480
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   8493
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4800
      Top             =   600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Database\Stu_Database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Database\Stu_Database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from BCA1st"
      Caption         =   "Adodc1"
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
      Left            =   9240
      TabIndex        =   19
      Top             =   5040
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line22 
      Visible         =   0   'False
      X1              =   5040
      X2              =   5040
      Y1              =   2520
      Y2              =   2880
   End
   Begin VB.Line Line21 
      Visible         =   0   'False
      X1              =   5160
      X2              =   5160
      Y1              =   1800
      Y2              =   1200
   End
   Begin VB.Line Line20 
      Visible         =   0   'False
      X1              =   2640
      X2              =   2640
      Y1              =   1800
      Y2              =   1200
   End
   Begin VB.Line Line19 
      Visible         =   0   'False
      X1              =   3960
      X2              =   3960
      Y1              =   1440
      Y2              =   1200
   End
   Begin VB.Line Line18 
      Visible         =   0   'False
      X1              =   1560
      X2              =   1560
      Y1              =   1440
      Y2              =   1200
   End
   Begin VB.Line Line17 
      Visible         =   0   'False
      X1              =   1560
      X2              =   5175
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line16 
      Visible         =   0   'False
      X1              =   3000
      X2              =   3000
      Y1              =   960
      Y2              =   1200
   End
   Begin VB.Line Line15 
      Visible         =   0   'False
      X1              =   3840
      X2              =   3840
      Y1              =   2160
      Y2              =   2640
   End
   Begin VB.Line Line14 
      Visible         =   0   'False
      X1              =   4200
      X2              =   4200
      Y1              =   2160
      Y2              =   3240
   End
   Begin VB.Line Line13 
      Visible         =   0   'False
      X1              =   5400
      X2              =   5400
      Y1              =   2520
      Y2              =   3480
   End
   Begin VB.Line Line12 
      Visible         =   0   'False
      X1              =   3000
      X2              =   3000
      Y1              =   2520
      Y2              =   3480
   End
   Begin VB.Line Line11 
      Visible         =   0   'False
      X1              =   2640
      X2              =   2640
      Y1              =   2520
      Y2              =   2880
   End
   Begin VB.Line Line10 
      Visible         =   0   'False
      X1              =   5640
      X2              =   5640
      Y1              =   4080
      Y2              =   2040
   End
   Begin VB.Line Line9 
      Visible         =   0   'False
      X1              =   5520
      X2              =   5640
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line8 
      Visible         =   0   'False
      X1              =   4440
      X2              =   4440
      Y1              =   3840
      Y2              =   1800
   End
   Begin VB.Line Line7 
      Visible         =   0   'False
      X1              =   4320
      X2              =   4440
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line6 
      Visible         =   0   'False
      X1              =   3240
      X2              =   3240
      Y1              =   4080
      Y2              =   2160
   End
   Begin VB.Line Line5 
      Visible         =   0   'False
      X1              =   3120
      X2              =   3240
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line4 
      Visible         =   0   'False
      X1              =   2040
      X2              =   2040
      Y1              =   3720
      Y2              =   2040
   End
   Begin VB.Line Line3 
      Visible         =   0   'False
      X1              =   1920
      X2              =   2040
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   1680
      X2              =   1680
      Y1              =   2160
      Y2              =   3120
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   1320
      X2              =   1320
      Y1              =   2160
      Y2              =   2520
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      Height          =   4815
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()




Line16.Visible = True
Line17.Visible = True
Line18.Visible = True
Line19.Visible = True
Line20.Visible = True
Line21.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True

Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False

Line5.Visible = False
Line6.Visible = False
Line11.Visible = False
Line12.Visible = False
Command9.Visible = False
Command10.Visible = False
Command11.Visible = False

Line7.Visible = False
Line8.Visible = False
Line14.Visible = False
Line15.Visible = False
Command12.Visible = False
Command13.Visible = False
Command14.Visible = False

Line9.Visible = False
Line10.Visible = False
Line13.Visible = False
Line22.Visible = False
Command15.Visible = False
Command16.Visible = False
Command17.Visible = False


End Sub


Private Sub Command19_Click()
Me.Hide
Form12.Show


Line16.Visible = True
Line17.Visible = True
Line18.Visible = True
Line19.Visible = True
Line20.Visible = True
Line21.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True

Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False

Line5.Visible = False
Line6.Visible = False
Line11.Visible = False
Line12.Visible = False
Command9.Visible = False
Command10.Visible = False
Command11.Visible = False

Line7.Visible = False
Line8.Visible = False
Line14.Visible = False
Line15.Visible = False
Command12.Visible = False
Command13.Visible = False
Command14.Visible = False

Line9.Visible = False
Line10.Visible = False
Line13.Visible = False
Line22.Visible = False
Command15.Visible = False
Command16.Visible = False
Command17.Visible = False


DataGrid1.Visible = False
End Sub

Private Sub Command2_Click()
Line1.Visible = True
Line2.Visible = True
Line3.Visible = True
Line4.Visible = True
Command6.Visible = True
Command7.Visible = True
Command8.Visible = True

DataGrid1.Visible = False

Line20.Visible = False
Line19.Visible = False
Line21.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False

End Sub

Private Sub Command3_Click()
Line5.Visible = True
Line6.Visible = True
Line11.Visible = True
Line12.Visible = True
Command9.Visible = True
Command10.Visible = True
Command11.Visible = True

DataGrid1.Visible = False


Line19.Visible = False
Line18.Visible = False
Line21.Visible = False
Command2.Visible = False
Command4.Visible = False
Command5.Visible = False
End Sub

Private Sub Command4_Click()
Line7.Visible = True
Line8.Visible = True
Line14.Visible = True
Line15.Visible = True
Command12.Visible = True
Command13.Visible = True
Command14.Visible = True

DataGrid1.Visible = False


Line20.Visible = False
Line18.Visible = False
Line21.Visible = False
Command3.Visible = False
Command2.Visible = False
Command5.Visible = False
End Sub

Private Sub Command5_Click()
Line9.Visible = True
Line10.Visible = True
Line13.Visible = True
Line22.Visible = True
Command15.Visible = True
Command16.Visible = True
Command17.Visible = True

DataGrid1.Visible = False


Line20.Visible = False
Line19.Visible = False
Line18.Visible = False
Command3.Visible = False
Command4.Visible = False
Command2.Visible = False
End Sub

Private Sub Command6_Click()
Adodc1.RecordSource = "select * from BCA1st"
Adodc1.Refresh
Line2.Visible = False
Line4.Visible = False
Line3.Visible = False
Command7.Visible = False
Command8.Visible = False

DataGrid1.Visible = True
End Sub

Private Sub Command7_Click()
Adodc1.RecordSource = "select * from BCA2nd"
Adodc1.Refresh


Line1.Visible = False
Line4.Visible = False
Line3.Visible = False
Command6.Visible = False
Command8.Visible = False

DataGrid1.Visible = True

End Sub

Private Sub Command8_Click()
Adodc1.RecordSource = "select * from BCA3rd"
Adodc1.Refresh



Line2.Visible = False
Line1.Visible = False
Command6.Visible = False
Command7.Visible = False

DataGrid1.Visible = True

End Sub

Private Sub Command9_Click()
Adodc1.RecordSource = "select * from BSC1st"
Adodc1.Refresh

Line12.Visible = False
Line5.Visible = False
Line6.Visible = False
Command10.Visible = False
Command11.Visible = False

DataGrid1.Visible = True

End Sub

Private Sub Command10_Click()
Adodc1.RecordSource = "select * from BSC2nd"
Adodc1.Refresh

Line11.Visible = False
Line4.Visible = False
Line6.Visible = False
Command9.Visible = False
Command11.Visible = False


DataGrid1.Visible = True
End Sub


Private Sub Command11_Click()
Adodc1.RecordSource = "select * from BSC3rd"
Adodc1.Refresh


Line11.Visible = False
Line12.Visible = False
Command9.Visible = False
Command10.Visible = False


DataGrid1.Visible = True
End Sub

Private Sub Command12_Click()
Adodc1.RecordSource = "select * from BCOM1st"
Adodc1.Refresh


Line14.Visible = False
Line7.Visible = False
Line8.Visible = False
Command13.Visible = False
Command14.Visible = False


DataGrid1.Visible = True
End Sub

Private Sub Command13_Click()
Adodc1.RecordSource = "select * from BCOM2nd"
Adodc1.Refresh


Line15.Visible = False
Line7.Visible = False
Line8.Visible = False
Command12.Visible = False
Command14.Visible = False


DataGrid1.Visible = True
End Sub


Private Sub Command14_Click()
Adodc1.RecordSource = "select * from BCOM3rd"
Adodc1.Refresh



Line15.Visible = False
Line14.Visible = False
Command12.Visible = False
Command13.Visible = False

DataGrid1.Visible = True
End Sub

Private Sub Command15_Click()
Adodc1.RecordSource = "select * from BA1st"
Adodc1.Refresh




Line13.Visible = False
Line9.Visible = False
Line10.Visible = False
Command16.Visible = False
Command17.Visible = False


DataGrid1.Visible = True
End Sub


Private Sub Command16_Click()
Adodc1.RecordSource = "select * from BA2nd"
Adodc1.Refresh


Line22.Visible = False
Line9.Visible = False
Line10.Visible = False
Command15.Visible = False
Command17.Visible = False

DataGrid1.Visible = True
End Sub


Private Sub Command17_Click()
Adodc1.RecordSource = "select * from BA2nd"
Adodc1.Refresh




Line13.Visible = False
Line22.Visible = False
Command16.Visible = False
Command15.Visible = False


DataGrid1.Visible = True
End Sub



