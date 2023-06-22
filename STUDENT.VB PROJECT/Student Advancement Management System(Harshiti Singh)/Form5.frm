VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00000040&
   Caption         =   "Update Record"
   ClientHeight    =   10140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14520
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   10140
   ScaleWidth      =   14520
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form5.frx":DD1F
      Height          =   2535
      Left            =   10680
      TabIndex        =   25
      Top             =   2520
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4471
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
            LCID            =   2057
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
            LCID            =   2057
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
      Height          =   615
      Left            =   11040
      Top             =   6600
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00004040&
      Caption         =   "Frame2"
      Height          =   7575
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Width           =   9375
      Begin VB.TextBox Text10 
         DataField       =   "Years"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   24
         Text            =   "Text10"
         Top             =   5640
         Width           =   3255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6960
         TabIndex        =   22
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox Text9 
         DataField       =   "Stream"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   21
         Text            =   "Text9"
         Top             =   5040
         Width           =   3255
      End
      Begin VB.TextBox Text8 
         DataField       =   "DOB"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   20
         Text            =   "Text8"
         Top             =   4440
         Width           =   3255
      End
      Begin VB.TextBox Text7 
         DataField       =   "RollNo"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   17
         Text            =   "Text7"
         Top             =   3840
         Width           =   3255
      End
      Begin VB.TextBox Text6 
         DataField       =   "Contact"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   15
         Text            =   "Text6"
         Top             =   3240
         Width           =   3255
      End
      Begin VB.TextBox Text5 
         DataField       =   "Class"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   2640
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         DataField       =   "Mother"
         DataSource      =   "Adodc1"
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
         TabIndex        =   13
         Text            =   "Text4"
         Top             =   1920
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         DataField       =   "Father"
         DataSource      =   "Adodc1"
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
         TabIndex        =   12
         Text            =   "Text3"
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         DataField       =   "Name"
         DataSource      =   "Adodc1"
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
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   480
         Width           =   3255
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
         Height          =   975
         Left            =   6960
         TabIndex        =   10
         Top             =   3840
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Update"
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
         Left            =   6960
         TabIndex        =   9
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackColor       =   &H00008000&
         Caption         =   "Year"
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
         Left            =   240
         TabIndex        =   23
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00008000&
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
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00008000&
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
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H00008000&
         Caption         =   "Regestraion no"
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
         Left            =   240
         TabIndex        =   16
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00008000&
         Caption         =   "Contact "
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
         Left            =   240
         TabIndex        =   8
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00008000&
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
         Left            =   240
         TabIndex        =   7
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00008000&
         Caption         =   "Mother's name"
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
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008000&
         Caption         =   "Father's name"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
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
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00004000&
      Caption         =   "Enter Registration Number"
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
      Height          =   1335
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   9375
      Begin VB.CommandButton Command1 
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
         Height          =   855
         Left            =   5640
         TabIndex        =   2
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H0080FF80&
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
         Left            =   480
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   480
         Width           =   3615
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command4_Click()
Adodc1.Refresh
End Sub

Private Sub Command1_Click()
Adodc1.Recordset.Find "RollNo=" & Val(Text1.Text)
If Adodc1.Recordset.EOF Then
MsgBox " Records doesn't exist"
End If
End Sub

Private Sub Command2_Click()

Adodc1.Recordset.Update
MsgBox "Updated Record"
End Sub

Private Sub Command3_Click()
Adodc1.Refresh
Form5.Hide
End Sub

