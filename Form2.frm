VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "welcome"
   ClientHeight    =   12075
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   22800
   LinkTopic       =   "Form2"
   ScaleHeight     =   12075
   ScaleWidth      =   22800
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Admin Log In"
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
      Height          =   4695
      Left            =   10080
      TabIndex        =   3
      Top             =   4200
      Width           =   7335
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         IMEMode         =   3  'DISABLE
         Left            =   3360
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3360
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00000000&
         Caption         =   "Admin Log In"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1800
         TabIndex        =   4
         Top             =   3240
         Width           =   3375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   " Password"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         TabIndex        =   8
         Top             =   1920
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   " Username"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         TabIndex        =   7
         Top             =   720
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3615
      Left            =   840
      TabIndex        =   0
      Top             =   4200
      Width           =   8175
      Begin VB.CommandButton Command3 
         Caption         =   "Feedback"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4440
         TabIndex        =   9
         Top             =   2160
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF80FF&
         Caption         =   "Book Now"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4320
         MaskColor       =   &H0000FF00&
         TabIndex        =   1
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         Caption         =   "Give Feedback :--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   435
         Left            =   720
         TabIndex        =   10
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         Caption         =   "Book Your Trip :--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   2
         Top             =   960
         Width           =   3150
      End
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   360
      Top             =   120
      Width           =   17295
   End
   Begin VB.Image Image2 
      Height          =   12015
      Left            =   0
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   22815
   End
   Begin VB.Menu Home 
      Caption         =   "Home"
   End
   Begin VB.Menu Out 
      Caption         =   "Log Out"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset

Private Sub Command1_Click()
Form2.Hide
Form12.Show

End Sub

Private Sub Command2_Click()
rs.MoveFirst
Dim i As Integer
i = 0
While Not rs.EOF = True
  
  If rs.Fields(18).Value = Text1.Text And rs.Fields(19).Value = Text2.Text Then
  i = i + 1
  
 Text1.Text = ""
 Text2.Text = ""
 MsgBox ("You Have Succesfully Sign In...")
 Form2.Hide
 Form13.Show
 
  End If
rs.MoveNext

Wend


  If i = 0 Then
  
  MsgBox ("You have enter invalid Username or Password..!")
  Text1.SetFocus
  End If
  
End Sub


Private Sub Command3_Click()
Form2.Hide
Form7.Show
End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture("J:\vb project\p1.jpg")

Set db = OpenDatabase("J:\vb project\employee1.mdb")
Set rs = db.OpenRecordset("select * from emp1")

End Sub

Private Sub out_Click()
Form2.Hide
Form1.Show

End Sub
