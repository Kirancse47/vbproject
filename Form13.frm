VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form13"
   ClientHeight    =   12075
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   22800
   LinkTopic       =   "Form13"
   ScaleHeight     =   12075
   ScaleWidth      =   22800
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "View Feedback"
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
      Left            =   5040
      TabIndex        =   3
      Top             =   6960
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4095
      Left            =   600
      TabIndex        =   0
      Top             =   4200
      Width           =   11055
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Caption         =   "Tourist Data"
         DownPicture     =   "Form13.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6720
         MaskColor       =   &H00FF00FF&
         TabIndex        =   2
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF00FF&
         Caption         =   "Employee Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2760
         TabIndex        =   1
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Choose Database..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   480
         TabIndex        =   4
         Top             =   360
         Width           =   4530
      End
   End
   Begin VB.Image Image1 
      Height          =   3255
      Left            =   240
      Stretch         =   -1  'True
      Top             =   360
      Width           =   13455
   End
   Begin VB.Image Image2 
      Height          =   12135
      Left            =   0
      Picture         =   "Form13.frx":6EA907
      Stretch         =   -1  'True
      Top             =   0
      Width           =   21615
   End
   Begin VB.Menu Home 
      Caption         =   "Home"
   End
   Begin VB.Menu logout 
      Caption         =   "Log Out"
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form13.Hide
Form15.Show
End Sub

Private Sub Command2_Click()
Form13.Hide
Form6.Show
End Sub

Private Sub Command3_Click()
Form13.Hide
Form18.Show
End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture("J:\vb project\p1.jpg")
End Sub

Private Sub Home_Click()
Form13.Hide
Form2.Show

End Sub

Private Sub Logout_Click()
Form13.Hide
Form2.Show
End Sub
