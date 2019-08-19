VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   12075
   ClientLeft      =   225
   ClientTop       =   270
   ClientWidth     =   22800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   12075
   ScaleWidth      =   22800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2280
      Top             =   5640
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      Caption         =   "LOG IN"
      Height          =   615
      Left            =   20280
      MaskColor       =   &H00800000&
      TabIndex        =   5
      Top             =   360
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      Caption         =   "PACKAGES"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3495
      Left            =   17640
      TabIndex        =   1
      Top             =   8280
      Width           =   4095
      Begin VB.CommandButton Command4 
         Caption         =   " ECONOMIC"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         MaskColor       =   &H00FF0000&
         TabIndex        =   4
         Top             =   2640
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   " GOLD"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         MaskColor       =   &H00FF0000&
         TabIndex        =   3
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "EXPENSIVE"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         MaskColor       =   &H00FF0000&
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.Image Image4 
      Height          =   3015
      Left            =   960
      Picture         =   "Form1.frx":2279
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   4935
   End
   Begin VB.Image Image3 
      Height          =   3015
      Left            =   960
      Picture         =   "Form1.frx":4317
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "EXPERIENCE"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   1455
      Left            =   9960
      TabIndex        =   7
      Top             =   7920
      Width           =   6015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "GO BEYOND TRAVEL"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2415
      Left            =   9720
      TabIndex        =   6
      Top             =   6720
      Width           =   10335
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   960
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   6975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CAPTURE THE MOMENT"
      BeginProperty Font 
         Name            =   "Magneto"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   11385
   End
   Begin VB.Image Image2 
      Height          =   12135
      Left            =   0
      Picture         =   "Form1.frx":6451
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()


Form2.Show
Form1.Hide



End Sub

Private Sub Command2_Click()
Form3.Show
Form1.Hide

End Sub

Private Sub Command3_Click()
Form4.Show
Form1.Hide
End Sub

Private Sub Command4_Click()
Form5.Show
Form1.Hide
End Sub



Private Sub Timer1_Timer()
If Image3.Visible = True Then
Image4.Visible = True
Image3.Visible = False
ElseIf Image4.Visible = True Then
Image3.Visible = True
Image4.Visible = False
End If
End Sub
