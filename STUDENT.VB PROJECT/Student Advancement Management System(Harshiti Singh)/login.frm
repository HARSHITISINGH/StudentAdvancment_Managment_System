VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF00FF&
   Caption         =   "Student Advancement Management System"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17295
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "login.frx":0000
   ScaleHeight     =   10035
   ScaleWidth      =   17295
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\adv\progressManagement.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LoginDetail"
      Top             =   6840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7800
      TabIndex        =   5
      Top             =   6480
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000040C0&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      MaskColor       =   &H000080FF&
      TabIndex        =   4
      Top             =   6600
      Width           =   2895
   End
   Begin VB.TextBox password 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   615
      Left            =   7080
      TabIndex        =   3
      Text            =   "pass"
      Top             =   5040
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   615
      Left            =   7080
      TabIndex        =   1
      Text            =   "123"
      Top             =   4080
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404000&
      Caption         =   "login page"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1575
      Left            =   6720
      TabIndex        =   6
      Top             =   1080
      Width           =   6375
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4200
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   735
      Left            =   4200
      TabIndex        =   2
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label U 
      BackColor       =   &H00C0C0FF&
      Caption         =   "User ID:"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4200
      TabIndex        =   0
      Top             =   4080
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset

Private Sub Command1_Click()
Dim user As String
Dim pass As String
Dim userid As String
Dim userpass As String

k = Text1.Text
kk = password.Text
Set db = OpenDatabase("D:\Student Advancement Management System(Harshiti Singh)\progressManagement1.mdb")
Set rs = db.OpenRecordset("select * from LoginDetail where ID=" + "'" + k + "'")

If rs.EOF Then
MsgBox "fill Currect ID"
 
Else
Set rss = db.OpenRecordset("select * from LoginDetail where Password=" + "'" + kk + "'")
 If rss.EOF Then
 MsgBox "Fill Correct Password"
 Else
 Form2.Show
 Form1.Hide
 End If
 

End If


End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Text1.Text = ""
password.Text = ""
Set db = OpenDatabase("D:\Student Advancement Management System(Harshiti Singh)\progressManagement1.mdb")

End Sub

