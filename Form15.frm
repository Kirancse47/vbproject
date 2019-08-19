VERSION 5.00
Begin VB.Form Form15 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form15"
   ClientHeight    =   12030
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   21630
   LinkTopic       =   "Form15"
   ScaleHeight     =   12030
   ScaleWidth      =   21630
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
      Left            =   6960
      TabIndex        =   5
      Top             =   7680
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   2040
      TabIndex        =   1
      Top             =   5760
      Width           =   13215
      Begin VB.CommandButton Command3 
         Caption         =   "Remove Employee"
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
         TabIndex        =   4
         Top             =   240
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Search Employee"
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
      Begin VB.CommandButton Command1 
         BackColor       =   &H00000000&
         Caption         =   "Add New Employee"
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
         TabIndex        =   2
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Image Image1 
      Height          =   3495
      Left            =   240
      Stretch         =   -1  'True
      Top             =   480
      Width           =   15255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome To Employee Database..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   555
      Left            =   360
      TabIndex        =   0
      Top             =   4680
      Width           =   8010
   End
   Begin VB.Image Image2 
      Height          =   11535
      Left            =   0
      Picture         =   "Form15.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   21615
   End
   Begin VB.Menu Home 
      Caption         =   "Home"
   End
   Begin VB.Menu Files 
      Caption         =   "Files"
      Begin VB.Menu Addnew 
         Caption         =   "Add New"
      End
      Begin VB.Menu Search 
         Caption         =   "Search"
      End
      Begin VB.Menu Delete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu Logout 
      Caption         =   "Log Out"
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Addnew_Click()
Form15.Hide
Form16.Show
End Sub

Private Sub Command1_Click()
Form15.Hide
Form16.Show
End Sub

Private Sub Command2_Click()
Form15.Hide
Form17.Show
End Sub

Private Sub Command3_Click()
Form15.Hide
Form17.Show
End Sub

Private Sub Command4_Click()
Form15.Hide
Form13.Show
End Sub

Private Sub Delete_Click()
Form15.Hide
Form17.Show
End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture("J:\vb project\p1.jpg")
End Sub

Private Sub Home_Click()
Form15.Hide
Form2.Show
End Sub

Private Sub Logout_Click()
Form15.Hide
Form2.Show
End Sub

Private Sub Search_Click()
Form15.Hide
Form17.Show
End Sub
