VERSION 5.00
Begin VB.Form Form5 
   Caption         =   " Economic"
   ClientHeight    =   12075
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   22800
   LinkTopic       =   "Form5"
   ScaleHeight     =   12075
   ScaleWidth      =   22800
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   4815
      Left            =   13440
      TabIndex        =   9
      Top             =   3960
      Width           =   3855
      Begin VB.Label Label10 
         Caption         =   " You Can Go......."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   615
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "1) DIGHA"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   600
         TabIndex        =   12
         Top             =   1560
         Width           =   1350
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "2) SUNDARBAN"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   600
         TabIndex        =   11
         Top             =   2520
         Width           =   2325
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "3) BHUTAN"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   600
         TabIndex        =   10
         Top             =   3480
         Width           =   1665
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6615
      Left            =   360
      TabIndex        =   0
      Top             =   3000
      Width           =   12375
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "* YOU GET A TOUR BUS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   8
         Top             =   4200
         Width           =   2340
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "* RS 2,500 CREDIT LIMIT FOR FOOD OF HOTELS RESTAURENT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   7
         Top             =   4680
         Width           =   6015
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "* YOU CAN SEE AND VIEW THE PLACE OF INSTEREST"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   6
         Top             =   3720
         Width           =   5130
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   " * IN THIS PACKAGE YOU GET 4 DAYS AND 3 NIGHTS TOUR"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   5
         Top             =   3120
         Width           =   5715
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "* THE MINIMUM COST IS 15,000/-"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   4
         Top             =   2520
         Width           =   3330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "* YOU GET 2 OR 3 STAR HOTEL"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   3
         Top             =   1920
         Width           =   3060
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   ">> ECONOMIC PACKAGE FEATURES :-"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   5430
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "ECONOMIC : LOW RANGE AND AFFORDABLE  PACKAGE"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   11730
      End
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   240
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11655
   End
   Begin VB.Image Image2 
      Height          =   12135
      Left            =   0
      Picture         =   "Form5.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22815
   End
   Begin VB.Menu Home 
      Caption         =   "Home"
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.Picture = LoadPicture("J:\vb project\p1.jpg")
End Sub

Private Sub Home_Click()
Form5.Hide
Form1.Show
End Sub


