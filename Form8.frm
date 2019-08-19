VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form8 
   Caption         =   " new save"
   ClientHeight    =   12075
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   22800
   LinkTopic       =   "Form8"
   ScaleHeight     =   12075
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
      Left            =   16800
      TabIndex        =   39
      Top             =   9600
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   15720
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Generate Sl.No"
      Height          =   495
      Left            =   7080
      TabIndex        =   38
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PHOTO"
      Height          =   435
      Left            =   12600
      TabIndex        =   37
      Top             =   4680
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
      Left            =   16800
      TabIndex        =   36
      Top             =   8400
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
      Left            =   16680
      TabIndex        =   35
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Payment Details"
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
      Left            =   720
      TabIndex        =   23
      Top             =   7680
      Width           =   8535
      Begin VB.TextBox Text15 
         Height          =   495
         Left            =   3000
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   2520
         Width           =   4695
      End
      Begin VB.TextBox Text14 
         Height          =   495
         Left            =   3000
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   1800
         Width           =   4695
      End
      Begin VB.TextBox Text13 
         Height          =   495
         Left            =   3000
         TabIndex        =   26
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
         TabIndex        =   29
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
         TabIndex        =   27
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
         TabIndex        =   25
         Top             =   480
         Width           =   1620
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Package Details"
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
      Left            =   9480
      TabIndex        =   13
      Top             =   6240
      Width           =   6735
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2160
         TabIndex        =   32
         Text            =   "Select"
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox Text11 
         Height          =   495
         Left            =   2160
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   3600
         Width           =   4215
      End
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   2160
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2760
         Width           =   4215
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   2160
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2040
         Width           =   4215
      End
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   2160
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Dutation :"
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
         TabIndex        =   22
         Top             =   3720
         Width           =   1170
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Hotel :"
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
         TabIndex        =   20
         Top             =   2880
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Fees :"
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
         TabIndex        =   18
         Top             =   2040
         Width           =   795
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Dep. Date :"
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
         TabIndex        =   16
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Place :"
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
         TabIndex        =   14
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3360
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2280
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
      Height          =   4695
      Left            =   720
      TabIndex        =   0
      Top             =   2880
      Width           =   8535
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
         TabIndex        =   34
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
         TabIndex        =   33
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   735
         Left            =   2160
         TabIndex        =   12
         Text            =   "Text6"
         Top             =   3120
         Width           =   5895
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   2280
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2400
         Width           =   4215
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   2280
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   2280
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   2280
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   480
         Width           =   4215
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
         TabIndex        =   9
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
         TabIndex        =   7
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
         TabIndex        =   5
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
         TabIndex        =   3
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
         TabIndex        =   1
         Top             =   480
         Width           =   915
      End
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2895
      Left            =   11760
      Stretch         =   -1  'True
      Top             =   1440
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
      Left            =   960
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   12135
      Left            =   120
      Picture         =   "Form8.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22695
   End
   Begin VB.Menu Home 
      Caption         =   "Home"
   End
   Begin VB.Menu Files 
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
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim k As Integer



Private Sub Combo1_Click()

If Combo1.ListIndex = 0 Then
Text9.Text = "85000"
Text10.Text = "Raj Hotel"
Text11.Text = "11 dayes"
End If
If Combo1.ListIndex = 1 Then
Text9.Text = "82000"
Text10.Text = "Royal Palace"
Text11.Text = "10 dayes"
End If

If Combo1.ListIndex = 2 Then
Text9.Text = "80000"
Text10.Text = "Techno Hotel"
Text11.Text = "9 dayes"
End If

If Combo1.ListIndex = 3 Then
Text9.Text = "42000"
Text10.Text = "Darjeeling Hotel"
Text11.Text = "7 dayes"
End If

If Combo1.ListIndex = 4 Then
Text9.Text = "41000"
Text10.Text = "NB Palace"
Text11.Text = "6 dayes"
End If

If Combo1.ListIndex = 5 Then
Text9.Text = "40000"
Text10.Text = "Hill Palace"
Text11.Text = "6 dayes"
End If

If Combo1.ListIndex = 6 Then
Text9.Text = "16000"
Text10.Text = "KB Hotel"
Text11.Text = "4 dayes"
End If

If Combo1.ListIndex = 7 Then
Text9.Text = "16000"
Text10.Text = "MG Hotel"
Text11.Text = "3 dayes"
End If

If Combo1.ListIndex = 8 Then
Text9.Text = "15000"
Text10.Text = "JK Hotel"
Text11.Text = "3 dayes"
End If

End Sub

Private Sub Command1_Click()
CommonDialog1.ShowOpen
Image1.Picture = LoadPicture(CommonDialog1.FileName)

End Sub

Private Sub Command2_Click()
Set db = OpenDatabase("J:\vb project\employee.mdb")
Set rs = db.OpenRecordset("select * from emp")

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

rs.Fields(7).Value = Combo1.Text
rs.Fields(8).Value = Text8.Text
rs.Fields(9).Value = CDbl(Text9.Text)
rs.Fields(10).Value = Text10.Text
rs.Fields(11).Value = Text11.Text
rs.Fields(12).Value = Text12.Text
rs.Fields(13).Value = CDbl(Text13.Text)
rs.Fields(14).Value = CStr(Text14.Text)
rs.Fields(15).Value = Text15.Text
rs.Fields(16).Value = CommonDialog1.FileName
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
Combo1.Text = "Select"
Image1.Picture = LoadPicture("J:\vb project\P3.JPG")


End Sub

Private Sub Command4_Click()
Set db = OpenDatabase("J:\vb project\employee.mdb")
Set rs = db.OpenRecordset("select * from emp")
If rs.EOF Then
Text1.Text = 1
Else
rs.MoveLast
k = rs.Fields(0).Value
Text1.Text = k + 1

End If




End Sub

Private Sub Command5_Click()
Form8.Hide
Form6.Show
End Sub

Private Sub Form_Load()
Combo1.AddItem "Kashmir(-Expensive)"
Combo1.AddItem "Goa(-Expensive)"
Combo1.AddItem "Rajashtan(-Expensive)"
Combo1.AddItem "Darjeeling(-Gold)"
Combo1.AddItem "Ghuahati(-Gold)"
Combo1.AddItem "Odisha(-Gold)"
Combo1.AddItem "Digha(-Economic)"
Combo1.AddItem "Sundarban(-Economic)"
Combo1.AddItem "Bhutan(-Economic)"

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


End Sub


Private Sub Home_Click()
Form8.Hide
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
Combo1.Text = "Select"
Image1.Picture = LoadPicture("J:\vb project\P3.JPG")
End Sub

Private Sub out_Click()
Form8.Hide
Form1.Show

End Sub
Private Sub Delete_Click()
Form9.Hide
Form10.Show
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
Combo1.Text = "Select"
Image1.Picture = LoadPicture("J:\vb project\P3.JPG")
Form8.Hide
Form9.Show

End Sub
