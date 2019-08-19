VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form17 
   Caption         =   "Form17"
   ClientHeight    =   12015
   ClientLeft      =   630
   ClientTop       =   510
   ClientWidth     =   22800
   LinkTopic       =   "Form17"
   ScaleHeight     =   12015
   ScaleWidth      =   22800
   StartUpPosition =   2  'CenterScreen
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
      Left            =   16560
      TabIndex        =   46
      Top             =   9240
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Next"
      Connect         =   "Access"
      DatabaseName    =   "J:\vb project\employee1.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   16080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "emp1"
      Top             =   10200
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   16440
      TabIndex        =   45
      Top             =   7920
      Width           =   1575
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
      TabIndex        =   29
      Top             =   1440
      Width           =   8535
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   2280
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   2280
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   2280
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   2280
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   2400
         Width           =   4215
      End
      Begin VB.TextBox Text6 
         Height          =   735
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   33
         Text            =   "Form17.frx":0000
         Top             =   3120
         Width           =   5895
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
         TabIndex        =   32
         Top             =   4080
         Width           =   1575
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
         TabIndex        =   31
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         Height          =   735
         Left            =   2160
         TabIndex        =   30
         Text            =   "Text7"
         Top             =   4800
         Width           =   5055
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
         TabIndex        =   43
         Top             =   480
         Width           =   915
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
         TabIndex        =   42
         Top             =   1080
         Width           =   645
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
         TabIndex        =   41
         Top             =   1800
         Width           =   1470
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
         TabIndex        =   40
         Top             =   2520
         Width           =   1455
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
         TabIndex        =   39
         Top             =   3120
         Width           =   1200
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
         TabIndex        =   38
         Top             =   4920
         Width           =   1155
      End
   End
   Begin VB.TextBox Text1 
      DataField       =   "slno"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2640
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   840
      Width           =   2655
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
      TabIndex        =   17
      Top             =   4320
      Width           =   6735
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   2160
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   2160
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   2160
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   2040
         Width           =   4215
      End
      Begin VB.TextBox Text11 
         Height          =   495
         Left            =   2280
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2880
         Width           =   4215
      End
      Begin VB.TextBox Text16 
         Height          =   495
         Left            =   2280
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   3600
         Width           =   4215
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
         TabIndex        =   27
         Top             =   600
         Width           =   690
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
         TabIndex        =   26
         Top             =   1320
         Width           =   1395
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
         TabIndex        =   25
         Top             =   2040
         Width           =   915
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
         TabIndex        =   24
         Top             =   2880
         Width           =   1620
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
         TabIndex        =   23
         Top             =   3600
         Width           =   1755
      End
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
      TabIndex        =   8
      Top             =   7800
      Width           =   8535
      Begin VB.TextBox Text12 
         Height          =   495
         Left            =   3000
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   480
         Width           =   4695
      End
      Begin VB.TextBox Text13 
         Height          =   495
         Left            =   3000
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1080
         Width           =   4695
      End
      Begin VB.TextBox Text14 
         Height          =   495
         Left            =   3000
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1800
         Width           =   4695
      End
      Begin VB.TextBox Text15 
         Height          =   495
         Left            =   3000
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   2520
         Width           =   4695
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
         TabIndex        =   16
         Top             =   480
         Width           =   1620
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
         TabIndex        =   15
         Top             =   1200
         Width           =   1065
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
         TabIndex        =   14
         Top             =   1920
         Width           =   765
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
         TabIndex        =   13
         Top             =   2640
         Width           =   2130
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   16440
      TabIndex        =   7
      Top             =   5520
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
      Height          =   855
      Left            =   16440
      TabIndex        =   6
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PHOTO"
      Height          =   435
      Left            =   11880
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Log In Details"
      Height          =   2175
      Left            =   9120
      TabIndex        =   0
      Top             =   9120
      Width           =   6495
      Begin VB.TextBox Text17 
         Height          =   495
         Left            =   2040
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox Text18 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1200
         Width           =   4215
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
         TabIndex        =   4
         Top             =   480
         Width           =   1455
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
         TabIndex        =   3
         Top             =   1320
         Width           =   1380
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   15600
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      TabIndex        =   44
      Top             =   840
      Width           =   1455
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
   Begin VB.Image Image2 
      Height          =   11895
      Left            =   120
      Picture         =   "Form17.frx":0006
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
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim db As Database
Dim rs As Recordset
Dim k As Integer

Private Sub Addnew_Click()
Form17.Hide
Form16.Show

End Sub

Private Sub Command2_Click()
k = CInt(Text1.Text)

Set db = OpenDatabase("J:\vb project\employee1.mdb")
Set rs = db.OpenRecordset("select * from emp1 where slno = " + Str(k))
If rs.EOF() Then
MsgBox ("Record not found...!")
Else
MsgBox ("Record found..\n click ok to show record")

Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
Text5.Text = rs.Fields(4).Value
Text6.Text = rs.Fields(5).Value
Text7.Text = rs.Fields(7).Value
Text8.Text = rs.Fields(8).Value
Text9.Text = rs.Fields(9).Value
Text10.Text = rs.Fields(10).Value
Text11.Text = rs.Fields(11).Value
Text12.Text = rs.Fields(13).Value
Text13.Text = rs.Fields(14).Value
Text14.Text = rs.Fields(15).Value
Text15.Text = rs.Fields(16).Value
Text16.Text = rs.Fields(12).Value
Text17.Text = rs.Fields(18).Value
Text18.Text = rs.Fields(19).Value

Image1.Picture = LoadPicture(rs.Fields(17))

If rs.Fields(6) = "Male" Then
Option1.Value = True
End If
If rs.Fields(6) = "Female" Then
Option2.Value = True
End If
End If
End Sub

Private Sub Command4_Click()
k = CInt(Text1.Text)

Set db = OpenDatabase("J:\vb project\employee1.mdb")
Set rs = db.OpenRecordset("select * from emp1 where slno = " + Str(k))
If rs.EOF() Then
MsgBox ("Record not found...!")
Else
rs.Delete
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
Image1.Picture = LoadPicture("D:\vb project\p3.JPG")

MsgBox ("Record Deleted...")
End If
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
Image1.Picture = LoadPicture("J:\vb project\p3.JPG")
End Sub

Private Sub Command5_Click()
Form17.Hide
Form15.Show
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
k = CInt(Text1.Text)

Set db = OpenDatabase("J:\vb project\employee1.mdb")
Set rs = db.OpenRecordset("select * from emp1 where slno = " + Str(k))
If rs.EOF() Then
MsgBox ("Record not found...!")
Else

Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
Text5.Text = rs.Fields(4).Value
Text6.Text = rs.Fields(5).Value
Text7.Text = rs.Fields(7).Value
Text8.Text = rs.Fields(8).Value
Text9.Text = rs.Fields(9).Value
Text10.Text = rs.Fields(10).Value
Text11.Text = rs.Fields(11).Value
Text12.Text = rs.Fields(13).Value
Text13.Text = rs.Fields(14).Value
Text14.Text = rs.Fields(15).Value
Text15.Text = rs.Fields(16).Value
Text16.Text = rs.Fields(12).Value
Text17.Text = rs.Fields(18).Value
Text18.Text = rs.Fields(19).Value

Image1.Picture = LoadPicture(rs.Fields(17))

If rs.Fields(6) = "Male" Then
Option1.Value = True
End If
If rs.Fields(6) = "Female" Then
Option2.Value = True
End If
End If
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
Image1.Picture = LoadPicture("J:\vb project\p3.JPG")

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
Image1.Picture = LoadPicture("J:\vb project\p3.JPG")


End Sub

Private Sub Home_Click()
Form17.Hide
Form2.Show
End Sub

Private Sub Logout_Click()
Form17.Hide
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
Image1.Picture = LoadPicture("J:\vb project\p3.JPG")

End Sub
