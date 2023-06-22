VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form18 
   BackColor       =   &H00004080&
   Caption         =   "Form18"
   ClientHeight    =   10155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17595
   LinkTopic       =   "Form18"
   ScaleHeight     =   10155
   ScaleWidth      =   17595
   StartUpPosition =   3  'Windows Default
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5775
      Left            =   1560
      OleObjectBlob   =   "FORM18.frx":0000
      TabIndex        =   10
      Top             =   3600
      Width           =   14535
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
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Text            =   "Text4"
      Top             =   1440
      Width           =   2055
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
      Left            =   13800
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   2160
      Width           =   1695
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
      Left            =   7440
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
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
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
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
      Left            =   2520
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Student Name"
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
      Left            =   360
      TabIndex        =   8
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
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
      Left            =   11880
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
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
      Left            =   5400
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Registration No."
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
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Student Progress Chart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   6720
      TabIndex        =   2
      Top             =   360
      Width           =   4815
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim db As Database
Dim rs As Recordset
Dim ss As Recordset
Dim k As String


Private Sub Command1_Click()
k = Text1.Text
Set db = OpenDatabase("D:\adv\progressManagement.mdb")
Set rs = db.OpenRecordset("select * from StudentRecord where RollNo=" + "'" + k + "'")

If rs.EOF() Then
MsgBox ("Record does not found")
Else
Text4.Text = rs.Fields(1)
Text2.Text = rs.Fields(14)
Text3.Text = rs.Fields(20)
Dim X(1 To 8, 1 To 6) As Variant
X(1, 2) = "Periodic Test-1"
X(1, 3) = "Periodic Test-2"

X(2, 1) = rs.Fields(27)
X(2, 2) = rs.Fields(9)
X(2, 3) = rs.Fields(15)

X(3, 1) = rs.Fields(28)
X(3, 2) = rs.Fields(10)
X(3, 3) = rs.Fields(16)

X(4, 1) = rs.Fields(29)
X(4, 2) = rs.Fields(11)
X(4, 3) = rs.Fields(17)

X(5, 1) = rs.Fields(30)
X(5, 2) = rs.Fields(12)
X(5, 3) = rs.Fields(18)

X(6, 1) = rs.Fields(31)
X(6, 2) = rs.Fields(13)
X(6, 3) = rs.Fields(19)







MSChart1.ChartData = X

End If


End Sub

Private Sub Form_Load()
MSChart1.Title = "Student Progress Chart of Periodic Test - 1 and Periodic Test - 2)"
MSChart1.ShowLegend = True


Dim X(1 To 8, 1 To 6) As Variant
X(1, 2) = "Periodic Test-1"
X(1, 3) = "Periodic Test-2"


MSChart1.ChartData = X
End Sub

