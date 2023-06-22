VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000040&
   Caption         =   "Search Record"
   ClientHeight    =   9405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16935
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   9405
   ScaleWidth      =   16935
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H000080FF&
      Caption         =   "Student Record"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   5775
      Left            =   7680
      TabIndex        =   3
      Top             =   4320
      Width           =   10095
      Begin VB.PictureBox Adodc1 
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   600
         ScaleHeight     =   555
         ScaleWidth      =   2115
         TabIndex        =   15
         Top             =   4320
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00404040&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3600
         TabIndex        =   14
         Top             =   3960
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         DataField       =   "Name"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3000
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   480
         Width           =   4335
      End
      Begin VB.TextBox Text6 
         DataField       =   "Contact"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         TabIndex        =   12
         Text            =   "Text6"
         Top             =   3240
         Width           =   4215
      End
      Begin VB.TextBox Text5 
         DataField       =   "Class"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   2520
         Width           =   4215
      End
      Begin VB.TextBox Text4 
         DataField       =   "Mother"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   10
         Text            =   "Text4"
         Top             =   1920
         Width           =   4335
      End
      Begin VB.TextBox Text3 
         DataField       =   "Father"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3000
         TabIndex        =   9
         Text            =   "Text3"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "Contact"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Caption         =   "Class"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "Mother's Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "Father 's Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Search Record"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   2175
      Left            =   7560
      TabIndex        =   0
      Top             =   2280
      Width           =   9615
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6600
         MaskColor       =   &H00004080&
         TabIndex        =   2
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   960
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   5415
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim db As Database
Dim rs As Recordset
Dim k As String
Private Sub Command1_Click()
k = Text1.Text
Set db = OpenDatabase("D:\adv\progressManagement.mdb")
Set rs = db.OpenRecordset("select * from StudentRecord where RollNo=" + "'" + k + "'")
If rs.EOF() Then
MsgBox ("Record does not found")
Else
Text2.Text = rs.Fields(1).Value()
Text3.Text = rs.Fields(2).Value()
Text4.Text = rs.Fields(3).Value()
Text5.Text = rs.Fields(4).Value()
Text6.Text = rs.Fields(5).Value()

End If
End Sub

Private Sub Command2_Click()
Form4.Hide

End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub

