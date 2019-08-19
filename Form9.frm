VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form9 
   Caption         =   "tsearch"
   ClientHeight    =   12075
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   22800
   LinkTopic       =   "Form9"
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
      Left            =   16200
      TabIndex        =   41
      Top             =   8280
      Width           =   1335
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
      Height          =   615
      Left            =   16200
      TabIndex        =   40
      Top             =   7320
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog3 
      Left            =   15720
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Data Data1 
      Caption         =   "Next"
      Connect         =   "Access"
      DatabaseName    =   "J:\vb project\employee.mdb"
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
      Height          =   975
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "emp"
      Top             =   9240
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PHOTO"
      Height          =   375
      Left            =   12360
      TabIndex        =   37
      Top             =   3240
      Width           =   1695
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
      Left            =   120
      TabIndex        =   23
      Top             =   600
      Width           =   8535
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   2280
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   2280
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   2280
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   2280
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   2400
         Width           =   4215
      End
      Begin VB.TextBox Text6 
         Height          =   735
         Left            =   2160
         TabIndex        =   26
         Text            =   "Text6"
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   4080
         Width           =   1695
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   3120
         Width           =   1200
      End
   End
   Begin VB.TextBox Text1 
      DataField       =   "ID"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2640
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   0
      Width           =   2655
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
      Left            =   8760
      TabIndex        =   11
      Top             =   3960
      Width           =   6735
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   2160
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   2160
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   2040
         Width           =   4215
      End
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   2160
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   2760
         Width           =   4215
      End
      Begin VB.TextBox Text11 
         Height          =   495
         Left            =   2160
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   3600
         Width           =   4215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2160
         TabIndex        =   12
         Text            =   "Select"
         Top             =   600
         Width           =   4215
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
         TabIndex        =   21
         Top             =   600
         Width           =   855
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
         TabIndex        =   20
         Top             =   1320
         Width           =   1395
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
         TabIndex        =   19
         Top             =   2040
         Width           =   795
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
         TabIndex        =   18
         Top             =   2880
         Width           =   810
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
         TabIndex        =   17
         Top             =   3720
         Width           =   1170
      End
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
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   8535
      Begin VB.TextBox Text12 
         Height          =   495
         Left            =   3000
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   480
         Width           =   4695
      End
      Begin VB.TextBox Text13 
         Height          =   495
         Left            =   3000
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1080
         Width           =   4695
      End
      Begin VB.TextBox Text14 
         Height          =   495
         Left            =   3000
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1800
         Width           =   4695
      End
      Begin VB.TextBox Text15 
         Height          =   495
         Left            =   3000
         TabIndex        =   3
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   2640
         Width           =   2130
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   " Search"
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
      TabIndex        =   1
      Top             =   4800
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
      TabIndex        =   0
      Top             =   6120
      Width           =   1455
   End
   Begin VB.PictureBox CommonDialog1 
      Height          =   480
      Left            =   4440
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   38
      Top             =   1200
      Width           =   1200
   End
   Begin VB.PictureBox CommonDialog2 
      Height          =   480
      Left            =   4440
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   39
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Image Image2 
      Height          =   2415
      Left            =   11760
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2895
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2895
      Index           =   1
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2895
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2895
      Index           =   0
      Left            =   0
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
      Top             =   0
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   12135
      Left            =   0
      Picture         =   "Form9.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22815
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
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim k As Integer

Private Sub Command2_Click()
k = CInt(Text1.Text)

Set db = OpenDatabase("J:\vb project\employee.mdb")
Set rs = db.OpenRecordset("select * from emp where ID = " + Str(k))
If rs.EOF() Then
MsgBox ("Record not found...!")
Else
MsgBox ("Record found..\n click ok to show record")

Text2.Text = rs.Fields(1).Value

Text3.Text = rs.Fields(2).Value

Text4.Text = rs.Fields(3).Value

Text5.Text = rs.Fields(4).Value

Text6.Text = rs.Fields(5).Value

Text8.Text = rs.Fields(8).Value

Text9.Text = rs.Fields(9).Value

Text10.Text = rs.Fields(10).Value

Text11.Text = rs.Fields(11).Value

Text12.Text = rs.Fields(12).Value
Text13.Text = rs.Fields(13).Value
Text14.Text = rs.Fields(14).Value
Text15.Text = rs.Fields(15).Value

Image2.Picture = LoadPicture(rs.Fields(16))


If rs.Fields(6) = "Male" Then
Option1.Value = True
End If

If rs.Fields(6) = "Female" Then
Option2.Value = True
End If


Combo1.Text = rs.Fields(7).Value

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
Combo1.Text = "Select"
Image2.Picture = LoadPicture("J:\vb project\p3.JPG")

End Sub

Private Sub Command4_Click()
k = CInt(Text1.Text)

Set db = OpenDatabase("J:\vb project\employee.mdb")
Set rs = db.OpenRecordset("select * from emp where ID = " + Str(k))
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
Combo1.Text = "Select"
Image2.Picture = LoadPicture("J:\vb project\p3.JPG")

MsgBox ("Record Deleted...")
End If
End Sub

Private Sub Command5_Click()
Form9.Hide
Form6.Show
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
Image2.Picture = LoadPicture("J:\vb project\P3.JPG")
k = CInt(Text1.Text)

Set db = OpenDatabase("J:\vb project\employee.mdb")
Set rs = db.OpenRecordset("select * from emp where ID = " + Str(k))
If rs.EOF() Then
MsgBox ("Record not found...!")
Else
Text2.Text = rs.Fields(1).Value

Text3.Text = rs.Fields(2).Value

Text4.Text = rs.Fields(3).Value

Text5.Text = rs.Fields(4).Value

Text6.Text = rs.Fields(5).Value

Text8.Text = rs.Fields(8).Value

Text9.Text = rs.Fields(9).Value

Text10.Text = rs.Fields(10).Value

Text11.Text = rs.Fields(11).Value

Text12.Text = rs.Fields(12).Value
Text13.Text = rs.Fields(13).Value
Text14.Text = rs.Fields(14).Value
Text15.Text = rs.Fields(15).Value

Image2.Picture = LoadPicture(rs.Fields(16))


If rs.Fields(6) = "Male" Then
Option1.Value = True
End If

If rs.Fields(6) = "Female" Then
Option2.Value = True
End If


Combo1.Text = rs.Fields(7).Value

End If


End Sub

Private Sub Delete_Click()
Form9.Hide
Form10.Show

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
Image2.Picture = LoadPicture("J:\vb project\P3.JPG")
Form9.Hide
Form8.Show

End Sub

Private Sub out_Click()
Form8.Hide
Form1.Show

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
Image2.Picture = LoadPicture("J:\vb project\P3.JPG")
End Sub

