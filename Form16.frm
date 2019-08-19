VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form16 
   Caption         =   "Form16"
   ClientHeight    =   12030
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   22800
   LinkTopic       =   "Form16"
   ScaleHeight     =   12030
   ScaleWidth      =   22800
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command5 
      Caption         =   "Home"
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
      Left            =   16200
      TabIndex        =   46
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Log In Details"
      Height          =   2175
      Left            =   9120
      TabIndex        =   41
      Top             =   9120
      Width           =   6495
      Begin VB.TextBox Text18 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox Text17 
         Height          =   495
         Left            =   2040
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   45
         Top             =   1320
         Width           =   1380
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Username :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   43
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Generate Sl.No"
      Height          =   495
      Left            =   6360
      TabIndex        =   35
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PHOTO"
      Height          =   435
      Left            =   11880
      TabIndex        =   34
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16080
      TabIndex        =   33
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15960
      TabIndex        =   32
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Bank Details"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   23
      Top             =   7800
      Width           =   8535
      Begin VB.TextBox Text15 
         Height          =   495
         Left            =   3000
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   2520
         Width           =   4695
      End
      Begin VB.TextBox Text14 
         Height          =   495
         Left            =   3000
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   1800
         Width           =   4695
      End
      Begin VB.TextBox Text13 
         Height          =   495
         Left            =   3000
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1080
         Width           =   4695
      End
      Begin VB.TextBox Text12 
         Height          =   495
         Left            =   3000
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Account Holder :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   31
         Top             =   2640
         Width           =   2130
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "IFSC :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   30
         Top             =   1920
         Width           =   765
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "A/C No :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   29
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Bank Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   28
         Top             =   480
         Width           =   1620
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Position Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   9000
      TabIndex        =   14
      Top             =   4320
      Width           =   6735
      Begin VB.TextBox Text16 
         Height          =   495
         Left            =   2280
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   3600
         Width           =   4215
      End
      Begin VB.TextBox Text11 
         Height          =   495
         Left            =   2280
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   2880
         Width           =   4215
      End
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   2160
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2040
         Width           =   4215
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   2160
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   2160
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Prev. work at :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   40
         Top             =   3600
         Width           =   1755
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Experience :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   22
         Top             =   2880
         Width           =   1620
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Salary :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   21
         Top             =   2040
         Width           =   915
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Join. Date :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   20
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Post :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   19
         Top             =   600
         Width           =   690
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2640
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   840
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Personal Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   8535
      Begin VB.TextBox Text7 
         Height          =   735
         Left            =   2160
         TabIndex        =   38
         Text            =   "Text7"
         Top             =   4800
         Width           =   5055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   4080
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   " Male"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   735
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "Form16.frx":0000
         Top             =   3120
         Width           =   5895
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   2280
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2400
         Width           =   4215
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   2280
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   2280
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   2280
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Email Id :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   37
         Top             =   4920
         Width           =   1155
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   12
         Top             =   3120
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Addhar No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   11
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Contact No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   10
         Top             =   1800
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "C/O :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   9
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   915
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   15600
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2895
      Left            =   11040
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Serial No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   240
      TabIndex        =   36
      Top             =   840
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   12015
      Left            =   120
      Picture         =   "Form16.frx":0006
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22695
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
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim k As Integer

Private Sub Addnew_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Image1.Picture = LoadPicture("J:\vb project\P3.JPG")

End Sub

Private Sub Command1_Click()
CommonDialog1.ShowOpen
Image1.Picture = LoadPicture(CommonDialog1.FileName)

End Sub

Private Sub Command2_Click()
Set db = OpenDatabase("J:\vb project\employee1.mdb")
Set rs = db.OpenRecordset("select * from emp1")

rs.Addnew

rs.Fields(0).Value = CInt(Text1.Text)
rs.Fields(1).Value = Text2.Text
rs.Fields(2).Value = Text3.Text
rs.Fields(3).Value = CDbl(Text4.Text)
rs.Fields(4).Value = CDbl(Text5.Text)
rs.Fields(5).Value = Text6.Text

If Option1.Value = True Then
rs.Fields(6) = "Male"
End If
If Option2.Value = True Then
rs.Fields(6) = "Female"
End If

rs.Fields(7).Value = Text7.Text
rs.Fields(8).Value = Text8.Text
rs.Fields(9).Value = Text9.Text
rs.Fields(10).Value = CDbl(Text10.Text)
rs.Fields(11).Value = Text11.Text
rs.Fields(12).Value = Text16.Text
rs.Fields(13).Value = Text12.Text
rs.Fields(14).Value = CDbl(Text13.Text)
rs.Fields(15).Value = CStr(Text14.Text)
rs.Fields(16).Value = Text15.Text
rs.Fields(17).Value = CommonDialog1.FileName
rs.Fields(18).Value = Text17.Text
rs.Fields(19).Value = Text18.Text
MsgBox ("Record saved")
rs.Update

End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Image1.Picture = LoadPicture("J:\vb project\P3.JPG")


End Sub

Private Sub Command4_Click()
Set db = OpenDatabase("J:\vb project\employee1.mdb")
Set rs = db.OpenRecordset("select * from emp1")
If rs.EOF Then
Text1.Text = 1
Else
rs.MoveLast
k = rs.Fields(0).Value
Text1.Text = k + 1

End If




End Sub

Private Sub Command5_Click()
Form17.Hide
Form15.Show
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
End Sub


Private Sub Home_Click()
Form16.Hide
Form2.Show

End Sub

Private Sub new_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Image1.Picture = LoadPicture("D:\vb project\P3.JPG")
End Sub

Private Sub out_Click()
Form8.Hide
Form1.Show

End Sub
Private Sub Delete_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Image1.Picture = LoadPicture("J:\vb project\P3.JPG")
Form16.Hide
Form17.Show
End Sub

Private Sub Logout_Click()
Form16.Hide
Form2.Show
End Sub

Private Sub Search_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Image1.Picture = LoadPicture("J:\vb project\P3.JPG")
Form16.Hide
Form17.Show

End Sub

