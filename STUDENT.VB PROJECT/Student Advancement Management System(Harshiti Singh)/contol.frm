VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00400040&
   Caption         =   "Form2"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19710
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   22.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "contol.frx":0000
   ScaleHeight     =   12375
   ScaleWidth      =   19710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command17 
      Caption         =   "Show Progress PT-1, PT-2 and Final"
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
      Left            =   8160
      TabIndex        =   22
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Show Progress PT-1 Vs PT-2"
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
      Left            =   4440
      TabIndex        =   21
      Top             =   5880
      Width           =   2775
   End
   Begin VB.CommandButton Command11 
      Caption         =   "LogOut"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6840
      TabIndex        =   18
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00400040&
      Caption         =   "MARK"
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
      Height          =   1575
      Left            =   3720
      TabIndex        =   15
      Top             =   4080
      Width           =   8295
      Begin VB.CommandButton Command15 
         Caption         =   "Add Marks-Final"
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
         Left            =   5640
         TabIndex        =   20
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Add Marks-PT-2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   3120
         TabIndex        =   17
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Add marks-PT-1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   360
         TabIndex        =   16
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00400040&
      Caption         =   "Exam"
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
      Height          =   2055
      Left            =   12960
      TabIndex        =   13
      Top             =   7440
      Width           =   3375
      Begin VB.CommandButton Command12 
         Caption         =   "update exam"
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
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Add Exam"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00004080&
      Caption         =   "Students Advancement Management System"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   2055
      Left            =   2520
      TabIndex        =   12
      Top             =   720
      Width           =   13695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00400000&
      Caption         =   "Years / Courses"
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
      Height          =   3735
      Left            =   13080
      TabIndex        =   7
      Top             =   3600
      Width           =   3255
      Begin VB.CommandButton Command9 
         Caption         =   "Update Stream"
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
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   1935
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Add Stream"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         Caption         =   "update year"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Add Year"
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
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400040&
      Caption         =   "Subject Panel"
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
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   6840
      Width           =   2895
      Begin VB.CommandButton Command5 
         Caption         =   "Update Subject"
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
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Add Subject"
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
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update Record"
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
      Left            =   360
      TabIndex        =   2
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Record"
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
      Left            =   360
      TabIndex        =   1
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400040&
      Caption         =   "Students Panel"
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
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   3015
      Begin VB.CommandButton Command3 
         Caption         =   "Search Record"
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
         Left            =   240
         MaskColor       =   &H0080FF80&
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Top             =   2400
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Show

End Sub



Private Sub Command10_Click()
Form13.Show

End Sub

Private Sub Command11_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub Command12_Click()
Form14.Show

End Sub

Private Sub Command13_Click()
Form15.Show

End Sub

Private Sub Command14_Click()
Form16.Show
End Sub

Private Sub Command15_Click()
Form17.Show
End Sub

Private Sub Command16_Click()
Form18.Show
End Sub

Private Sub Command17_Click()
Form19.Show
End Sub

Private Sub Command2_Click()
Form5.Show
End Sub
Private Sub Command3_Click()
Form4.Show
End Sub


Private Sub Command4_Click()
subject.Show

End Sub

Private Sub Command5_Click()
Form9.Show
End Sub

Private Sub Command6_Click()

Form6.Show
End Sub

Private Sub Command7_Click()
Form7.Show
End Sub

Private Sub Command8_Click()
Form8.Show
End Sub

Private Sub Command9_Click()
Form10.Show
End Sub

