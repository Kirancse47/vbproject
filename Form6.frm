VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFF80&
   Caption         =   " Tourist Details"
   ClientHeight    =   12075
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   22800
   LinkTopic       =   "Form6"
   ScaleHeight     =   12075
   ScaleWidth      =   22800
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Back"
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
      Left            =   7200
      TabIndex        =   5
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   1920
      TabIndex        =   1
      Top             =   5040
      Width           =   13215
      Begin VB.CommandButton Command1 
         BackColor       =   &H00000000&
         Caption         =   "Add New Booking"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Search Booking"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4920
         TabIndex        =   3
         Top             =   240
         Width           =   3015
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Remove Booking"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8520
         TabIndex        =   2
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   " Welcome to Tourist's Database......"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   555
      Left            =   600
      TabIndex        =   0
      Top             =   3600
      Width           =   8190
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   360
      Top             =   120
      Width           =   17775
   End
   Begin VB.Image Image2 
      Height          =   12135
      Left            =   0
      Picture         =   "Form6.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22815
   End
   Begin VB.Menu Home 
      Caption         =   "Home"
   End
   Begin VB.Menu files 
      Caption         =   "Files"
      Begin VB.Menu new 
         Caption         =   "New Record"
      End
      Begin VB.Menu search 
         Caption         =   "Search"
      End
      Begin VB.Menu delete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu out 
      Caption         =   "Log Out"
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form6.Hide
Form8.Show
End Sub

Private Sub Command2_Click()
Form6.Hide
Form9.Show
End Sub

Private Sub Command3_Click()
Form6.Hide
Form10.Show

End Sub

Private Sub Command4_Click()
Form6.Hide
Form13.Show

End Sub

Private Sub Delete_Click()
Form6.Hide
Form10.Show

End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture("J:\vb project\p1.jpg")
End Sub

Private Sub new_Click()
Form6.Hide
Form8.Show

End Sub

Private Sub Home_Click()
Form6.Hide
Form2.Show

End Sub

Private Sub out_Click()
Form6.Hide
Form1.Show

End Sub

Private Sub Search_Click()
Form6.Hide
Form9.Show

End Sub
