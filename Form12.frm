VERSION 5.00
Begin VB.Form Form12 
   Caption         =   "Form12"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   LinkTopic       =   "Form12"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Customer Sign In"
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
      Height          =   6375
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   9135
      Begin VB.CommandButton Command4 
         Caption         =   "Cancel"
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
         Left            =   3000
         TabIndex        =   10
         Top             =   5400
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3360
         TabIndex        =   7
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
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
         IMEMode         =   3  'DISABLE
         Left            =   3360
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   2400
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Reset"
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
         Left            =   4560
         TabIndex        =   5
         Top             =   4320
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sign In"
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
         Left            =   1440
         TabIndex        =   4
         Top             =   4320
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   " Username"
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
         Left            =   480
         TabIndex        =   9
         Top             =   1200
         Width           =   1905
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   "Password"
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
         Left            =   480
         TabIndex        =   8
         Top             =   2640
         Width           =   1740
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF80&
      Caption         =   " "
      Height          =   4455
      Left            =   10560
      TabIndex        =   0
      Top             =   1680
      Width           =   6135
      Begin VB.CommandButton Command2 
         Caption         =   "Sign Up"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   1680
         TabIndex        =   2
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "For Customer Sign Up..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   540
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   5160
      End
   End
   Begin VB.Image Image1 
      Height          =   12375
      Left            =   120
      Picture         =   "Form12.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22575
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset



Private Sub Command1_Click()
rs.MoveFirst
Dim i As Integer
i = 0
While Not rs.EOF = True
  
  If rs.Fields(0).Value = Text1.Text And rs.Fields(1).Value = Text2.Text Then
  i = i + 1
  
 Text1.Text = ""
 Text2.Text = ""
 MsgBox ("You Have Succesfully Sign In...")
 Form12.Hide
 Form14.Show
 
  End If
rs.MoveNext

Wend


  If i = 0 Then
  
  MsgBox ("You have enter invalid Username or Password..!")
  Text1.SetFocus
  End If
  
End Sub

Private Sub Command2_Click()
Form11.Show
Form12.Hide

End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Command4_Click()
Form2.Show
Form12.Hide

End Sub

Private Sub Form_Load()
Set db = OpenDatabase("J:\vb project\signup.mdb")
Set rs = db.OpenRecordset("select * from up")

End Sub
