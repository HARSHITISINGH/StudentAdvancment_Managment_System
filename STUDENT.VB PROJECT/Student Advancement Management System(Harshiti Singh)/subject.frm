VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form subject 
   BackColor       =   &H00C0FFFF&
   Caption         =   "subject"
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11145
   FillColor       =   &H00C0E0FF&
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form6"
   ScaleHeight     =   8280
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   4560
      Top             =   4800
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
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
      RecordSource    =   "Subject"
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
   Begin VB.Frame Frame1 
      Caption         =   "SUBJECT NAME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3615
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   8055
      Begin VB.CommandButton Command5 
         BackColor       =   &H0000FF00&
         Caption         =   "EXIT"
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
         Left            =   5160
         MaskColor       =   &H000080FF&
         TabIndex        =   5
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FF00&
         Caption         =   "CLEAR"
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
         Left            =   2760
         MaskColor       =   &H000080FF&
         TabIndex        =   4
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "ADD"
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
         Left            =   480
         MaskColor       =   &H000080FF&
         TabIndex        =   3
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         DataField       =   "Subject_NAme"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1080
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1560
         Width           =   4695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "SUBJECT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   3495
      End
   End
End
Attribute VB_Name = "subject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command4_Click()
Text1.Text = ""
End Sub

Private Sub Command5_Click()
subject.Hide
End Sub

