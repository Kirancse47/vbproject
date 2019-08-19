VERSION 5.00
Begin VB.Form Form18 
   Caption         =   "Form18"
   ClientHeight    =   12060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   LinkTopic       =   "Form18"
   ScaleHeight     =   12060
   ScaleWidth      =   22800
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Customer Feedback"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   9135
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   11535
      Begin VB.Data Data1 
         Caption         =   "Next"
         Connect         =   "Access"
         DatabaseName    =   "J:\vb project\employee.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
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
         Left            =   4200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "feedback"
         Top             =   7440
         Width           =   3975
      End
      Begin VB.TextBox Text5 
         DataField       =   "cmnt"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   4080
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   4920
         Width           =   4935
      End
      Begin VB.TextBox Text4 
         DataField       =   "rate"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   4080
         TabIndex        =   5
         Top             =   3960
         Width           =   4215
      End
      Begin VB.TextBox Text3 
         DataField       =   "email"
         DataSource      =   "Data1"
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
         Left            =   4080
         TabIndex        =   4
         Top             =   3000
         Width           =   4215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Back"
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
         Left            =   720
         TabIndex        =   3
         Top             =   7440
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         DataField       =   "mobile"
         DataSource      =   "Data1"
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
         IMEMode         =   3  'DISABLE
         Left            =   4080
         TabIndex        =   2
         Top             =   2040
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         DataField       =   "name"
         DataSource      =   "Data1"
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
         TabIndex        =   1
         Top             =   960
         Width           =   4215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   "Email :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   840
         TabIndex        =   11
         Top             =   3120
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   "Mobile No :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   840
         TabIndex        =   10
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   "Name :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   840
         TabIndex        =   9
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   "Comment :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   840
         TabIndex        =   8
         Top             =   5280
         Width           =   2025
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   "Rate Out Of 10 :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   840
         TabIndex        =   7
         Top             =   4200
         Width           =   2925
      End
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   12000
      Top             =   720
      Width           =   5895
   End
   Begin VB.Image Image2 
      Height          =   10095
      Left            =   120
      Picture         =   "Form18.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   21495
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form18.Hide
Form13.Show

End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture("J:\vb project\p1.jpg")
End Sub
