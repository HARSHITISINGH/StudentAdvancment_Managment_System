VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form19 
   BackColor       =   &H00008000&
   Caption         =   "Form19"
   ClientHeight    =   10485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17655
   LinkTopic       =   "Form19"
   Picture         =   "Form19.frx":0000
   ScaleHeight     =   10485
   ScaleWidth      =   17655
   StartUpPosition =   3  'Windows Default
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5895
      Left            =   1080
      OleObjectBlob   =   "Form19.frx":1D0F
      TabIndex        =   12
      Top             =   3600
      Width           =   15135
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13800
      TabIndex        =   9
      Text            =   "Text4"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9480
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000A&
      Caption         =   "Student Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   10
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000A&
      Caption         =   "Final (Total)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12240
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000A&
      Caption         =   "PT-2 (Total)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000A&
      Caption         =   "PT-1 (Total)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Registration No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Student Progress Chart"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   6120
      TabIndex        =   2
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim db As Database
Dim rs As Recordset
Dim ss As Recordset

Private Sub Command1_Click()
k = Text1.Text
Set db = OpenDatabase("D:\adv\progressManagement.mdb")
Set rs = db.OpenRecordset("select * from StudentRecord where RollNo=" + "'" + k + "'")

If rs.EOF() Then
MsgBox ("Record does not found")
Else
Text5.Text = rs.Fields(1)
Text2.Text = rs.Fields(14)
Text3.Text = rs.Fields(20)
Text4.Text = rs.Fields(26)
Dim X(1 To 8, 1 To 6) As Variant
X(1, 2) = "Periodic Test-1"
X(1, 3) = "Periodic Test-2"
X(1, 4) = "Final Exam"

X(2, 1) = rs.Fields(27)
X(2, 2) = rs.Fields(9)
X(2, 3) = rs.Fields(15)
X(2, 4) = rs.Fields(21)

X(3, 1) = rs.Fields(28)
X(3, 2) = rs.Fields(10)
X(3, 3) = rs.Fields(16)
X(3, 4) = rs.Fields(22)

X(4, 1) = rs.Fields(29)
X(4, 2) = rs.Fields(11)
X(4, 3) = rs.Fields(17)
X(4, 4) = rs.Fields(23)

X(5, 1) = rs.Fields(30)
X(5, 2) = rs.Fields(12)
X(5, 3) = rs.Fields(18)
X(5, 4) = rs.Fields(24)

X(6, 1) = rs.Fields(31)
X(6, 2) = rs.Fields(13)
X(6, 3) = rs.Fields(19)
X(6, 4) = rs.Fields(25)






MSChart1.ChartData = X

End If


End Sub

Private Sub Form_Load()
MSChart1.Title = "Student Progress Chart of Periodic Test - 1, Periodic Test - 2 and Final Exam)"
MSChart1.ShowLegend = True


Dim X(1 To 8, 1 To 6) As Variant
X(1, 2) = "Periodic Test-1"
X(1, 3) = "Periodic Test-2"
X(1, 4) = "Final Exam"

MSChart1.ChartData = X
End Sub

