VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00404080&
   Caption         =   "EXAMUPDTE"
   ClientHeight    =   10110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17070
   LinkTopic       =   "Form12"
   ScaleHeight     =   10110
   ScaleWidth      =   17070
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   13920
      ScaleHeight     =   270
      ScaleWidth      =   2355
      TabIndex        =   7
      Top             =   3480
      Width           =   2415
   End
   Begin VB.PictureBox DataGrid1 
      Height          =   1215
      Left            =   9840
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00004000&
      Caption         =   "UPDATE EXAM"
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
      Height          =   6735
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   11295
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
         Height          =   615
         Left            =   7560
         TabIndex        =   5
         Top             =   3600
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
         Height          =   615
         Left            =   4920
         TabIndex        =   4
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "UPDATE"
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
         Left            =   1920
         TabIndex        =   3
         Top             =   3600
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
         Height          =   615
         Left            =   3240
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2040
         Width           =   4455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         Caption         =   "UPDATE EXAM"
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
         Left            =   3360
         TabIndex        =   1
         Top             =   960
         Width           =   3375
      End
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

