VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H0080FFFF&
   Caption         =   "end"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   LinkTopic       =   "Form10"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Give Feedback"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      TabIndex        =   3
      Top             =   8640
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12360
      TabIndex        =   2
      Top             =   8640
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   3615
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   5415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "For Booking..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   690
      Left            =   5400
      TabIndex        =   1
      Top             =   7320
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Thank You....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   690
      Left            =   1680
      TabIndex        =   0
      Top             =   6120
      Width           =   3930
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox ("Thank You... Visit Again")
Form10.Hide
Form2.Show

End Sub

Private Sub Command2_Click()
Form10.Hide
Form7.Show
End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture("J:\vb project\p6.jpg")
End Sub

