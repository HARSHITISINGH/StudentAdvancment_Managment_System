VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form16 
   BackColor       =   &H0000C000&
   Caption         =   "Form16"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18300
   LinkTopic       =   "Form16"
   ScaleHeight     =   10215
   ScaleWidth      =   18300
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Student Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   360
      TabIndex        =   17
      Top             =   240
      Width           =   15615
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form16.frx":0000
         Height          =   2655
         Left            =   11760
         TabIndex        =   46
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   4683
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
         Height          =   375
         Left            =   2520
         Top             =   2880
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
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
      Begin VB.TextBox Text1 
         DataField       =   "RollNo"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3120
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         DataField       =   "Name"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3120
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         DataField       =   "Father"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3120
         TabIndex        =   22
         Text            =   "Text3"
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         DataField       =   "DOB"
         DataSource      =   "Adodc1"
         Height          =   525
         Left            =   9120
         TabIndex        =   21
         Text            =   "Text4"
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         DataField       =   "Class"
         DataSource      =   "Adodc1"
         Height          =   615
         Left            =   9120
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox Text11 
         DataField       =   "Stream"
         DataSource      =   "Adodc1"
         Height          =   525
         Left            =   9120
         TabIndex        =   19
         Text            =   "Text11"
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox Text12 
         DataField       =   "Years"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   9120
         TabIndex        =   18
         Text            =   "Text12"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Registration / Roll No."
         BeginProperty Font 
            Name            =   "News706 BT"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   31
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Student Name"
         BeginProperty Font 
            Name            =   "News706 BT"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   30
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Father's Name"
         BeginProperty Font 
            Name            =   "News706 BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   29
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Stream"
         BeginProperty Font 
            Name            =   "News706 BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         TabIndex        =   28
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Class"
         BeginProperty Font 
            Name            =   "News706 BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7440
         TabIndex        =   27
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "DOB"
         BeginProperty Font 
            Name            =   "News706 BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7560
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "News706 BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         TabIndex        =   25
         Top             =   1800
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Add Marks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   720
      TabIndex        =   2
      Top             =   3840
      Width           =   15375
      Begin VB.TextBox Text19 
         DataField       =   "PT_2_per"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "News706 BT"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8040
         TabIndex        =   48
         Text            =   "Text19"
         Top             =   4320
         Width           =   1695
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   45
         Text            =   "Combo5"
         Top             =   4560
         Width           =   1335
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   44
         Text            =   "Combo4"
         Top             =   3840
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   43
         Text            =   "Combo3"
         Top             =   3120
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   42
         Text            =   "Combo2"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   41
         Text            =   "Combo1"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "News706 BT"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   10560
         TabIndex        =   40
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox Text18 
         DataField       =   "PT2_T"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8040
         TabIndex        =   39
         Text            =   "Text18"
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "News706 BT"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   10560
         TabIndex        =   37
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "News706 BT"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   10560
         TabIndex        =   36
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox Text13 
         DataField       =   "PT2_S_1"
         DataSource      =   "Adodc1"
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
         Left            =   5280
         TabIndex        =   12
         Text            =   "Text13"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text14 
         DataField       =   "PT2_S_2"
         DataSource      =   "Adodc1"
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
         Left            =   5280
         TabIndex        =   11
         Text            =   "Text14"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox Text15 
         DataField       =   "PT2_S_3"
         DataSource      =   "Adodc1"
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
         Left            =   5280
         TabIndex        =   10
         Text            =   "Text15"
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox Text16 
         DataField       =   "PT2_S_4"
         DataSource      =   "Adodc1"
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
         Left            =   5280
         TabIndex        =   9
         Text            =   "Text16"
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox Text17 
         DataField       =   "PT2_S_5"
         DataSource      =   "Adodc1"
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
         Left            =   5280
         TabIndex        =   8
         Text            =   "Text17"
         Top             =   4560
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         DataField       =   "Subject_1"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   7
         Text            =   "Text6"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox Text7 
         DataField       =   "Subject_2"
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
         Left            =   2280
         TabIndex        =   6
         Text            =   "Text7"
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Text8 
         DataField       =   "Subject_3"
         DataSource      =   "Adodc1"
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
         Left            =   2280
         TabIndex        =   5
         Text            =   "Text8"
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox Text9 
         DataField       =   "Subject_4"
         DataSource      =   "Adodc1"
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
         Left            =   2280
         TabIndex        =   4
         Text            =   "Text9"
         Top             =   3720
         Width           =   2055
      End
      Begin VB.TextBox Text10 
         DataField       =   "Subject_5"
         DataSource      =   "Adodc1"
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
         Left            =   2280
         TabIndex        =   3
         Text            =   "Text10"
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label Label17 
         Caption         =   "Percent"
         BeginProperty Font 
            Name            =   "News706 BT"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8040
         TabIndex        =   47
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "News706 BT"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8040
         TabIndex        =   38
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "PT-2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   35
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Select Subject"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "Selected Subject"
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
         Left            =   2280
         TabIndex        =   33
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Marks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   32
         Top             =   600
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000001&
         FillColor       =   &H00C0FFFF&
         Height          =   855
         Left            =   0
         Top             =   360
         Width           =   15225
      End
      Begin VB.Label Label7 
         Caption         =   "PT-1"
         Height          =   375
         Left            =   8760
         TabIndex        =   16
         Top             =   480
         Width           =   1575
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         Height          =   3975
         Left            =   0
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Select Subject"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Selected Subject"
         Height          =   495
         Left            =   2520
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Marks"
         Height          =   375
         Left            =   5760
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD"
      Height          =   735
      Left            =   8280
      TabIndex        =   1
      Top             =   5280
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "EXIT"
      Height          =   735
      Left            =   8280
      TabIndex        =   0
      Top             =   7080
      Width           =   2895
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim per As Integer
Dim db As Database
Dim rs As Recordset

Private Sub Combo1_Click()
Text6.Text = Combo1.List(Combo1.ListIndex)
End Sub

Private Sub Combo2_Click()
Text7.Text = Combo2.List(Combo2.ListIndex)
End Sub

Private Sub Combo3_Click()
Text8.Text = Combo3.List(Combo3.ListIndex)
End Sub

Private Sub Combo4_Click()
Text9.Text = Combo4.List(Combo4.ListIndex)
End Sub

Private Sub Combo5_Click()
Text10.Text = Combo5.List(Combo5.ListIndex)
End Sub

Private Sub Command1_Click()
Adodc1.Recordset.Update
MsgBox "Marks has Recorded Sucessfully"
End Sub


Private Sub Command3_Click()
Form16.Hide
End Sub

Private Sub Command5_Click()
Text18.Text = Val(Text13.Text) + Val(Text14.Text) + Val(Text15.Text) + Val(Text16.Text) + Val(Text17.Text)
Text19.Text = Val(Text18.Text) / 5
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("D:\adv\progressManagement.mdb")
Set rs = db.OpenRecordset("select * from Subject where Subject_NAme is not null")
'Set ry = db.OpenRecordset("select Years from Years where Years is not null")

Do While Not rs.EOF

 Combo1.AddItem (rs.Fields(0))
 Combo2.AddItem (rs.Fields(0))
  Combo3.AddItem (rs.Fields(0))
   Combo4.AddItem (rs.Fields(0))
    Combo5.AddItem (rs.Fields(0))
rs.MoveNext

Loop


End Sub

