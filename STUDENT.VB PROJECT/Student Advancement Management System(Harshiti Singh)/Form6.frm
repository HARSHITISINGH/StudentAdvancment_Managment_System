VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   BackColor       =   &H0080FFFF&
   Caption         =   "YEAR "
   ClientHeight    =   10665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13380
   LinkTopic       =   "Form6"
   ScaleHeight     =   10665
   ScaleWidth      =   13380
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   615
      Left            =   4080
      Top             =   5400
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1085
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
      RecordSource    =   "Years"
      Caption         =   "Adodc2"
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
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9120
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400040&
      Caption         =   "Add Year"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   4815
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   12015
      Begin VB.CommandButton Command4 
         Caption         =   "delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7680
         TabIndex        =   6
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5520
         TabIndex        =   5
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   4
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   3
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         DataField       =   "Years"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2760
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Fill the Year"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   1
         Top             =   480
         Width           =   3615
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command2_Click()
Text1.Text = ""
End Sub

Private Sub Command3_Click()
Form6.Hide
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.Delete
End Sub

