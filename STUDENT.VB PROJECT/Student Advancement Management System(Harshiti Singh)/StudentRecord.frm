VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BackColor       =   &H0000FFFF&
   Caption         =   "StudentRecord"
   ClientHeight    =   9870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16740
   FillColor       =   &H0000FFFF&
   ForeColor       =   &H0080FFFF&
   LinkTopic       =   "Form3"
   Picture         =   "StudentRecord.frx":0000
   ScaleHeight     =   9870
   ScaleWidth      =   16740
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   480
      Top             =   9600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\adv\progressManagement.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\adv\progressManagement.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "StudentRecord"
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
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   14640
      TabIndex        =   25
      Text            =   "Combo2"
      Top             =   9360
      Width           =   1815
   End
   Begin VB.TextBox Text9 
      DataField       =   "Years"
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
      Left            =   7320
      TabIndex        =   23
      Text            =   "Text9"
      Top             =   8520
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   14520
      TabIndex        =   20
      Text            =   "Combo1"
      Top             =   8400
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      DataField       =   "Stream"
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
      Left            =   7320
      TabIndex        =   19
      Text            =   "Text8"
      Top             =   7920
      Width           =   3975
   End
   Begin VB.TextBox Text7 
      DataField       =   "DOB"
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
      Left            =   7320
      TabIndex        =   18
      Text            =   "Text7"
      Top             =   7320
      Width           =   3975
   End
   Begin VB.CommandButton Command4 
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
      Height          =   1095
      Left            =   12000
      TabIndex        =   15
      Top             =   5400
      Width           =   2535
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
      Height          =   735
      Left            =   7320
      TabIndex        =   14
      Text            =   "Text6"
      Top             =   6600
      Width           =   3975
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
      Height          =   735
      Left            =   7320
      TabIndex        =   12
      Text            =   "Text5"
      Top             =   5760
      Width           =   3975
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
      Height          =   735
      Left            =   7320
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   4920
      Width           =   3975
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
      Height          =   735
      Left            =   7320
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   4080
      Width           =   3975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12000
      MaskColor       =   &H00400000&
      TabIndex        =   6
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12000
      TabIndex        =   5
      Top             =   2760
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
      Height          =   735
      Left            =   7320
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   3240
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      DataField       =   "RollNo"
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
      Height          =   735
      Left            =   7320
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Add Student Records "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   5040
      TabIndex        =   0
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label Label11 
      BackColor       =   &H000080FF&
      Caption         =   "Years"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   24
      Top             =   9360
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H000080FF&
      Caption         =   "Years"
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
      Left            =   3840
      TabIndex        =   22
      Top             =   8640
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackColor       =   &H000080FF&
      Caption         =   "Select Stream"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      TabIndex        =   21
      Top             =   8400
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H000080FF&
      Caption         =   "Stream"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3840
      TabIndex        =   17
      Top             =   7920
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H000080FF&
      Caption         =   "DOB"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3840
      TabIndex        =   16
      Top             =   7200
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H000080FF&
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
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3840
      TabIndex        =   13
      Top             =   6480
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H000080FF&
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
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   3840
      TabIndex        =   11
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
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
      ForeColor       =   &H00000040&
      Height          =   735
      Left            =   3840
      TabIndex        =   9
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   "Father's Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   3840
      TabIndex        =   7
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
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
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   3840
      TabIndex        =   2
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "Registration Number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   3840
      TabIndex        =   1
      Top             =   2520
      Width           =   2415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim db As Database
Dim rs As Recordset
Dim ry As Recordset
Private Sub Combo1_Click()
Text8.Text = Combo1.List(Combo1.ListIndex)
End Sub

Private Sub Combo2_Click()
Text9.Text = Combo2.List(Combo2.ListIndex)
End Sub

Private Sub Command1_Click()
Adodc1.Recordset.AddNew

End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""

End Sub

Private Sub Command4_Click()
Form3.Hide
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Set db = OpenDatabase("D:\adv\progressManagement.mdb")
Set rs = db.OpenRecordset("select Stream_Name from Stream where Stream_Name is not null")
Set ry = db.OpenRecordset("select Years from Years where Years is not null")

Do While Not rs.EOF

 Combo1.AddItem (rs.Fields(0))

rs.MoveNext

Loop

Do While Not ry.EOF

 Combo2.AddItem (ry.Fields(0))

ry.MoveNext

Loop
End Sub

