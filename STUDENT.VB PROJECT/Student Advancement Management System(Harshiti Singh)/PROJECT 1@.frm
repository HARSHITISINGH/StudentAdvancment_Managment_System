VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00FF0000&
   Caption         =   "Update Subject"
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14925
   LinkTopic       =   "Form9"
   ScaleHeight     =   8280
   ScaleWidth      =   14925
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   11280
      ScaleHeight     =   675
      ScaleWidth      =   1275
      TabIndex        =   7
      Top             =   4680
      Width           =   1335
   End
   Begin VB.PictureBox DataGrid1 
      Height          =   1455
      Left            =   10320
      ScaleHeight     =   1395
      ScaleWidth      =   4275
      TabIndex        =   6
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Update Subject"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   9975
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0E0FF&
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
         Height          =   735
         Left            =   6600
         MaskColor       =   &H00C0C0FF&
         TabIndex        =   5
         Top             =   3000
         Width           =   1815
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
         Height          =   735
         Left            =   4080
         TabIndex        =   4
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
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
         Height          =   735
         Left            =   1440
         TabIndex        =   3
         Top             =   3000
         Width           =   1815
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
         Height          =   855
         Left            =   2040
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1680
         Width           =   4455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Update Subject"
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
         Left            =   2160
         TabIndex        =   1
         Top             =   600
         Width           =   3975
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Adodc1.Recordset.Update
End Sub

Private Sub Command2_Click()
Text1.Text = ""
End Sub

Private Sub Command3_Click()
Form9.Hide
End Sub
