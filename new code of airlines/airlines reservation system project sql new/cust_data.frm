VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form8 
   BackColor       =   &H80000005&
   Caption         =   "Customer Information"
   ClientHeight    =   6705
   ClientLeft      =   1620
   ClientTop       =   1710
   ClientWidth     =   9855
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9855
   Begin MSComctlLib.ImageList ImageList 
      Left            =   10200
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   87
      ImageHeight     =   26
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cust_data.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cust_data.frx":1B22
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cust_data.frx":4F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cust_data.frx":537C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cust_data.frx":8F2E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   240
      Picture         =   "cust_data.frx":93E9
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   120
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   7440
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      Caption         =   "Customer Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   240
      TabIndex        =   14
      Top             =   840
      Width           =   9495
      Begin VB.ComboBox title_box 
         Height          =   315
         ItemData        =   "cust_data.frx":9839
         Left            =   240
         List            =   "cust_data.frx":9846
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox fname_box 
         Height          =   375
         Left            =   1200
         MaxLength       =   19
         TabIndex        =   1
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox lname_box 
         Height          =   375
         Left            =   2880
         MaxLength       =   19
         TabIndex        =   2
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox birth_date 
         Height          =   315
         ItemData        =   "cust_data.frx":9857
         Left            =   4920
         List            =   "cust_data.frx":98B8
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1440
         Width           =   615
      End
      Begin VB.ComboBox birth_month 
         Height          =   315
         ItemData        =   "cust_data.frx":992F
         Left            =   5640
         List            =   "cust_data.frx":9957
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1440
         Width           =   615
      End
      Begin VB.ComboBox birth_year 
         Height          =   315
         ItemData        =   "cust_data.frx":999A
         Left            =   6240
         List            =   "cust_data.frx":9A5B
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox mobile_box 
         Height          =   375
         Left            =   7080
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox add1_box 
         Height          =   375
         Left            =   240
         MaxLength       =   49
         TabIndex        =   7
         Top             =   2640
         Width           =   3135
      End
      Begin VB.TextBox add2_box 
         Height          =   375
         Left            =   3960
         MaxLength       =   49
         TabIndex        =   8
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox city_box 
         Height          =   375
         Left            =   240
         MaxLength       =   10
         TabIndex        =   9
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox state_box 
         Height          =   375
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   10
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox country_box 
         Height          =   375
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   11
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox pin_box 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   375
         Left            =   6000
         MaxLength       =   8
         TabIndex        =   12
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton continue_booking_button 
         Height          =   495
         Left            =   6000
         Picture         =   "cust_data.frx":9BD9
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         Height          =   735
         Left            =   5880
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Image Image10 
         Height          =   570
         Left            =   7800
         Picture         =   "cust_data.frx":D77B
         Top             =   600
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.Image Image9 
         Enabled         =   0   'False
         Height          =   390
         Left            =   6840
         Picture         =   "cust_data.frx":10B65
         Top             =   3120
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Image Image8 
         Enabled         =   0   'False
         Height          =   390
         Left            =   4560
         Picture         =   "cust_data.frx":12677
         Top             =   3000
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Image Image7 
         Enabled         =   0   'False
         Height          =   390
         Left            =   2520
         Picture         =   "cust_data.frx":14189
         Top             =   3000
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Image Image6 
         Enabled         =   0   'False
         Height          =   390
         Left            =   720
         Picture         =   "cust_data.frx":15C9B
         Top             =   3000
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   480
         Picture         =   "cust_data.frx":177AD
         Top             =   600
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000005&
         Caption         =   "Title*"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000005&
         Caption         =   "First Name*"
         Height          =   375
         Left            =   1200
         TabIndex        =   24
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000005&
         Caption         =   "Last Name*"
         Height          =   255
         Left            =   2880
         TabIndex        =   23
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000005&
         Caption         =   "Date Of Birth"
         Height          =   375
         Left            =   4920
         TabIndex        =   22
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000005&
         Caption         =   "Mobile No*"
         Height          =   375
         Left            =   7080
         TabIndex        =   21
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000005&
         Caption         =   "Address-Line 1*"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000005&
         Caption         =   "Address-Line 2"
         Height          =   375
         Left            =   3960
         TabIndex        =   19
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000005&
         Caption         =   "City*"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000005&
         Caption         =   "State*"
         Height          =   375
         Left            =   2040
         TabIndex        =   17
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000005&
         Caption         =   "Country*"
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000005&
         Caption         =   "Pin Code*"
         Height          =   255
         Left            =   6000
         TabIndex        =   15
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image Image2 
         Enabled         =   0   'False
         Height          =   390
         Left            =   1920
         Picture         =   "cust_data.frx":192BF
         Top             =   600
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Image Image3 
         Enabled         =   0   'False
         Height          =   390
         Left            =   3720
         Picture         =   "cust_data.frx":1ADD1
         Top             =   600
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Image Image4 
         Enabled         =   0   'False
         Height          =   390
         Left            =   8040
         Picture         =   "cust_data.frx":1C8E3
         Top             =   600
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Image Image5 
         Enabled         =   0   'False
         Height          =   390
         Left            =   1560
         Picture         =   "cust_data.frx":1E3F5
         Top             =   2160
         Visible         =   0   'False
         Width           =   1305
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pnrno As String
Dim i As Integer
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub add1_box_Change()
Image5.Visible = False
End Sub
Private Sub city_box_Change()
Image6.Visible = False
End Sub
Private Sub Command1_Click()
Form3.Visible = True
db.Close
Unload Me
End Sub
Private Sub continue_booking_button_Click()
On Error GoTo err_h:
i = 0
Call check_all_fill
If i = 0 Then
pnrno = Left(fname_box.Text, 3) & Right(mobile_box.Text, 5) & CInt(Int(Rnd() * Int(Rnd() * 199)))
rs.Open "insert into cust_data values('" & title_box.Text & "','" & fname_box.Text & "','" & lname_box.Text & "','" & birth_date.Text & "','" & birth_month.Text & "','" & birth_year.Text & "','" & mobile_box.Text & "','" & add1_box.Text & "','" & add2_box.Text & "','" & city_box.Text & "','" & state_box.Text & "','" & country_box.Text & "','" & pin_box.Text & "','" & pnrno & "','" & flight_no & "','" & cabin & "')", db, adOpenDynamic, adLockOptimistic, adCmdText
first_name = fname_box.Text
last_name = lname_box.Text
pnr_no = pnrno
continue_booking_button.Picture = ImageList.ListImages(5).Picture
dup_ticket = "false"
Timer1.Enabled = True
End If
Exit Sub
err_h:
MsgBox "fill all the checkbox correctly", vbExclamation
End Sub
Private Sub country_box_Change()
Image8.Visible = False
End Sub
Private Sub fname_box_Change()
Image2.Visible = False
End Sub
Public Function check_all_fill()
If title_box.Text = "" Then
Image1.Visible = True
i = 1
End If
If fname_box.Text = "" Then
  Image2.Visible = True
i = 1
  End If
If lname_box.Text = "" Then
Image3.Visible = True
i = 1
End If
 
  If mobile_box.Text = "" Then
 Image4.Visible = True
 i = 1
 End If
 
 If add1_box.Text = "" Then
 Image5.Visible = True
 i = 1
 End If
 
 If city_box.Text = "" Then
 Image6.Visible = True
 i = 1
 End If
 
  If state_box.Text = "" Then
  Image7.Visible = True
  i = 1
  End If
  If country_box.Text = "" Then
  Image8.Visible = True
 i = 1
  End If
  If pin_box.Text = "" Then
  Image9.Visible = True
  i = 1
  End If
  
 If Image10.Visible = True Then
 i = 1
 End If
End Function
Private Sub Form_Load()
Image1.Picture = ImageList.ListImages(1).Picture
Image2.Picture = ImageList.ListImages(1).Picture
Image3.Picture = ImageList.ListImages(1).Picture
Image4.Picture = ImageList.ListImages(1).Picture
Image5.Picture = ImageList.ListImages(1).Picture
Image6.Picture = ImageList.ListImages(1).Picture
Image7.Picture = ImageList.ListImages(1).Picture
Image8.Picture = ImageList.ListImages(1).Picture
Image9.Picture = ImageList.ListImages(1).Picture
Image10.Picture = ImageList.ListImages(2).Picture
continue_booking_button.Picture = ImageList.ListImages(4).Picture
Command1.Picture = ImageList.ListImages(3).Picture
'db.ConnectionString = "dsn=airlines_data;uid=system;pwd=harish"
db.ConnectionString = "dsn=airlines_data;uid=" & user_name & ";pwd=" & pass_word & " "
db.Open
End Sub

Private Sub lname_box_Change()
Image3.Visible = False
End Sub

Private Sub mobile_box_Change()
Image4.Visible = False
Image10.Visible = False
End Sub


Private Sub mobile_box_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyBack Then
Exit Sub
End If
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
End Sub

Private Sub mobile_box_LostFocus()
If Len(mobile_box.Text) <= 5 Then
Image10.Visible = True
i = 1
Else
i = 0
End If
End Sub

Private Sub pin_box_Change()
Image9.Visible = False
End Sub

Private Sub pin_box_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyBack Then
Exit Sub
End If
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
End Sub

Private Sub state_box_Change()
Image7.Visible = False
End Sub

Private Sub Timer1_Timer()
continue_booking_button.Picture = ImageList.ListImages(4).Picture
db.Close
Form7.Visible = True
Form8.Visible = False
Timer1.Enabled = False
End Sub

Private Sub title_box_Change()
Image1.Visible = False
End Sub
Private Sub title_box_Click()
Image1.Visible = False
End Sub
