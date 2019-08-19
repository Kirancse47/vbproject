VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00FFFF80&
   Caption         =   "feedback"
   ClientHeight    =   12030
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   22800
   LinkTopic       =   "Form7"
   ScaleHeight     =   12030
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
      Height          =   8175
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      Begin VB.TextBox Text1 
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
         TabIndex        =   8
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox Text2 
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
         TabIndex        =   7
         Top             =   2040
         Width           =   4215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Submit"
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
         Left            =   3720
         TabIndex        =   6
         Top             =   7200
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
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
         Left            =   600
         TabIndex        =   5
         Top             =   7200
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
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
         Height          =   855
         Left            =   6960
         TabIndex        =   4
         Top             =   7200
         Width           =   2295
      End
      Begin VB.TextBox Text3 
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
         TabIndex        =   3
         Top             =   3000
         Width           =   4215
      End
      Begin VB.TextBox Text4 
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
         TabIndex        =   2
         Top             =   3960
         Width           =   4215
      End
      Begin VB.TextBox Text5 
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
         TabIndex        =   1
         Top             =   4920
         Width           =   4935
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
         TabIndex        =   13
         Top             =   4200
         Width           =   2925
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
         TabIndex        =   12
         Top             =   5280
         Width           =   2025
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
         TabIndex        =   11
         Top             =   1080
         Width           =   1395
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
         TabIndex        =   9
         Top             =   3120
         Width           =   1350
      End
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   12360
      Top             =   720
      Width           =   5895
   End
   Begin VB.Image Image2 
      Height          =   12015
      Left            =   120
      Picture         =   "Form7.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22695
   End
   Begin VB.Menu Home 
      Caption         =   "Home"
   End
   Begin VB.Menu out 
      Caption         =   "Log Out"
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
MsgBox ("Invalid Entery...Fill Up all the fields..")
Text1.SetFocus

Else

Set db = OpenDatabase("J:\vb project\employee.mdb")
Set rs = db.OpenRecordset("select * from feedback")

rs.Addnew

rs.Fields(0).Value = Text1.Text
rs.Fields(1).Value = CDbl(Text2.Text)
rs.Fields(2).Value = Text3.Text
rs.Fields(3).Value = CInt(Text4.Text)
rs.Fields(4).Value = Text5.Text


MsgBox ("You Succecefully Registered...")
rs.Update
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Form7.Hide
Form2.Show

End If

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub

Private Sub Command3_Click()
Form7.Hide
Form2.Show

End Sub


Private Sub Form_Load()
Image1.Picture = LoadPicture("J:\vb project\p1.jpg")
End Sub

Private Sub Home_Click()
Form7.Hide
Form2.Show

End Sub


Private Sub out_Click()
Form7.Hide
Form1.Show

End Sub

