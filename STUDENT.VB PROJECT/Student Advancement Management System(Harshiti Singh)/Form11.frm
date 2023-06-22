VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00004040&
   Caption         =   "EXAM"
   ClientHeight    =   10275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17100
   LinkTopic       =   "Form11"
   ScaleHeight     =   10275
   ScaleWidth      =   17100
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   11400
      ScaleHeight     =   555
      ScaleWidth      =   2595
      TabIndex        =   8
      Top             =   4680
      Width           =   2655
   End
   Begin VB.PictureBox DataGrid1 
      Height          =   1815
      Left            =   8520
      ScaleHeight     =   1755
      ScaleWidth      =   2595
      TabIndex        =   6
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000040&
      Caption         =   "ADD EXAM "
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
      Height          =   6975
      Left            =   720
      TabIndex        =   0
      Top             =   2400
      Width           =   10815
      Begin VB.CommandButton Command3 
         Caption         =   "EXIT"
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
         Left            =   5880
         TabIndex        =   5
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CLEAR"
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
         Left            =   3720
         TabIndex        =   4
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ADD"
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
         Left            =   840
         TabIndex        =   3
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2160
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1800
         Width           =   4215
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "ADD EXAM"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         TabIndex        =   1
         Top             =   600
         Width           =   4215
      End
   End
   Begin VB.OLE OLE1 
      Height          =   1215
      Left            =   4440
      TabIndex        =   7
      Top             =   3360
      Width           =   3015
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Adodc1.Recordset.AddNew
End Sub
 
 Text1.Text = ""

