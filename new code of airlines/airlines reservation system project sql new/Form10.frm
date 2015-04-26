VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form10 
   BackColor       =   &H80000005&
   Caption         =   "Database Tree view"
   ClientHeight    =   9300
   ClientLeft      =   1860
   ClientTop       =   1470
   ClientWidth     =   9675
   LinkTopic       =   "Form10"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   9675
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Flight Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5535
      Left            =   4680
      TabIndex        =   27
      Top             =   1920
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Image Image5 
         Height          =   255
         Left            =   4560
         ToolTipText     =   "Close"
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "Flight Company"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   2400
         TabIndex        =   47
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "Departure City"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   2400
         TabIndex        =   45
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000005&
         Caption         =   "Arrival City"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   2400
         TabIndex        =   43
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000005&
         Caption         =   "Departure Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   2400
         TabIndex        =   41
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000005&
         Caption         =   "Departure Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   40
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   2400
         TabIndex        =   39
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   3240
         TabIndex        =   38
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000005&
         Caption         =   "Remaining economy seat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   37
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   2400
         TabIndex        =   36
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000005&
         Caption         =   "Price of economy seat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   2400
         TabIndex        =   34
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000005&
         Caption         =   "Remaining business seat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   2400
         TabIndex        =   32
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000005&
         Caption         =   "Price of business seat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   2400
         TabIndex        =   30
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label20 
         BackColor       =   &H80000005&
         Caption         =   "Flight No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label Label21 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   2400
         TabIndex        =   28
         Top             =   4800
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      Caption         =   "Customer Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   7215
      Left            =   4680
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Image Image6 
         Height          =   255
         Left            =   4560
         ToolTipText     =   "Close"
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label23 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label24 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   24
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label25 
         BackColor       =   &H80000005&
         Caption         =   "DOB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label26 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   1440
         TabIndex        =   22
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label27 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   1920
         TabIndex        =   21
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label28 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   2400
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label29 
         BackColor       =   &H80000005&
         Caption         =   "Mobile no"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label30 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   1440
         TabIndex        =   18
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label31 
         BackColor       =   &H80000005&
         Caption         =   "Address Line one"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label32 
         BackColor       =   &H80000005&
         Height          =   735
         Left            =   1440
         TabIndex        =   16
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Label Label33 
         BackColor       =   &H80000005&
         Caption         =   "Address Line two"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label34 
         BackColor       =   &H80000005&
         Height          =   735
         Left            =   1440
         TabIndex        =   14
         Top             =   3120
         Width           =   3135
      End
      Begin VB.Label Label35 
         BackColor       =   &H80000005&
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label36 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   4200
         Width           =   2415
      End
      Begin VB.Label Label37 
         BackColor       =   &H80000005&
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   4680
         Width           =   735
      End
      Begin VB.Label Label38 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   4680
         Width           =   2415
      End
      Begin VB.Label Label39 
         BackColor       =   &H80000005&
         Caption         =   "Country"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label Label40 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   5160
         Width           =   2415
      End
      Begin VB.Label Label41 
         BackColor       =   &H80000005&
         Caption         =   "Pin Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   5640
         Width           =   975
      End
      Begin VB.Label Label42 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label Label43 
         BackColor       =   &H80000005&
         Caption         =   "PNR NO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   6120
         Width           =   735
      End
      Begin VB.Label Label44 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label45 
         BackColor       =   &H80000005&
         Caption         =   "FLIGHT NO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   6720
         Width           =   1095
      End
      Begin VB.Label Label46 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   6720
         Width           =   1335
      End
   End
   Begin VB.CommandButton search 
      BackColor       =   &H80000005&
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   4320
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000005&
      Caption         =   "Seach PNRNO"
      Height          =   375
      Left            =   6240
      TabIndex        =   56
      Top             =   3840
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000005&
      Caption         =   "Search by flight no"
      Height          =   375
      Left            =   6240
      TabIndex        =   55
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6240
      TabIndex        =   54
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000005&
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7575
      Left            =   0
      TabIndex        =   49
      Top             =   1560
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   13361
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   0
      MousePointer    =   4
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8520
      Top             =   9720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form10.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form10.frx":0712
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form10.frx":2264
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form10.frx":2C56
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form10.frx":37CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form10.frx":3DE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form10.frx":440E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form10.frx":5A68
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form10.frx":5EB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form10.frx":71B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form10.frx":9632
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image8 
      Height          =   1920
      Left            =   7800
      Top             =   0
      Width           =   1920
   End
   Begin VB.Image Image7 
      Height          =   345
      Left            =   8160
      Picture         =   "Form10.frx":15684
      Top             =   2520
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label50 
      BackColor       =   &H80000005&
      Caption         =   "Passanger  PNR no"
      Height          =   255
      Left            =   6120
      TabIndex        =   53
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label49 
      BackColor       =   &H80000005&
      Caption         =   "Flight no"
      Height          =   255
      Left            =   6120
      TabIndex        =   52
      Top             =   120
      Width           =   855
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   5160
      Top             =   720
      Width           =   735
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   5160
      Top             =   120
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   2280
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label48 
      BackColor       =   &H80000005&
      Caption         =   "Departure to arrival city"
      Height          =   255
      Left            =   3000
      TabIndex        =   51
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label47 
      BackColor       =   &H80000005&
      Caption         =   "Departure date"
      Height          =   375
      Left            =   3000
      TabIndex        =   50
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   2280
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim rs2 As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs As New ADODB.Recordset




Private Sub Command1_Click()
db.Close
Unload Form10

End Sub

Private Sub Form_Click()
'Frame1.Visible = False
'Frame2.Visible = False
Image7.Visible = False
End Sub

Private Sub Form_Load()
On Error GoTo err_h:
db.ConnectionString = "dsn=airlines_data;uid=" & user_name & ";pwd=" & pass_word & " "
'db.ConnectionString = "dsn=airlines_data;uid=system;pwd=harish"

db.Open
rs.Open "select * from flight_data", db, adOpenDynamic, adLockOptimistic, adCmdText
rs1.Open "select * from flight_data", db, adOpenDynamic, adLockOptimistic, adCmdText
rs2.Open "select * from cust_data", db, adOpenDynamic, adLockOptimistic, adCmdText


Command1.Picture = ImageList1.ListImages(7).Picture
Image1.Picture = ImageList1.ListImages(1).Picture
Image2.Picture = ImageList1.ListImages(3).Picture
Image3.Picture = ImageList1.ListImages(4).Picture
Image4.Picture = ImageList1.ListImages(6).Picture
Image5.Picture = ImageList1.ListImages(8).Picture
Image6.Picture = ImageList1.ListImages(8).Picture
Image7.Picture = ImageList1.ListImages(9).Picture
search.Picture = ImageList1.ListImages(10).Picture
Image8.Picture = ImageList1.ListImages(11).Picture
Call treeview_data
Exit Sub
err_h:
MsgBox "error occureed while connecting to the database", vbInformation
Unload Form10


End Sub

Public Function treeview_data()
On Error GoTo err_h:
Dim i As Integer, j As Integer, k As Integer, l As Integer, m As Integer
Dim Data As Node
Set Data = TreeView1.Nodes.Add
Data.Image = 2
TreeView1.Nodes(1).Text = "Flight"
i = 0
Do While (Not rs.EOF)
i = i + 1
rs.MoveNext
Loop
rs.MoveFirst

'Text1.Text = i

Dim node1() As Node
ReDim node1(i) As Node
Dim temparray() As String
ReDim temparray(i) As String
m = 0
For j = 1 To i
    For k = 0 To m
      If (temparray(k) = rs.Fields(3)) Then
      l = 1
      End If
     Next k
      If l = 0 Then
      Set node1(j) = TreeView1.Nodes.Add(Data, tvwChild, , rs.Fields(3))
      'TreeView1.node1(j).Image = 1
      node1(j).Image = 1
      temparray(m) = rs.Fields(3)
      m = m + 1
      End If
    rs.MoveNext
    l = 0
Next j

'TreeView1.Nodes(1).Sorted = True
TreeView1.Nodes.Item(2).Sorted = True
'***************************************************************************
rs.MoveFirst

Dim children1 As Integer
children1 = TreeView1.Nodes(1).Children
Dim node2() As Node
ReDim node2(children1) As Node
Dim node3 As Node
Dim node4 As Node
i = 0
Do While (Not rs.EOF)
i = i + 1
rs.MoveNext
Loop


Dim z As Integer
rs.MoveFirst

For z = 1 To children1
rs.MoveFirst
ReDim temparray(i) As String
  m = 0
  l = 0
  For j = 1 To i

     For k = 0 To m 'problem change to i
      If (temparray(k) = (rs.Fields(1) + " " + "To" + " " + rs.Fields(2))) Then
      l = 1
      End If
     Next k
     'MsgBox TreeView1.Nodes.Item(z + 1)
     If l = 0 And (rs.Fields(3) = TreeView1.Nodes.Item(z + 1)) Then
     
      Set node2(z) = TreeView1.Nodes.Add(TreeView1.Nodes.Item(z + 1), tvwChild, , rs.Fields(1) + " " + "To" + " " + rs.Fields(2))
       node2(z).Image = 3
      temparray(m) = rs.Fields(1) + " " + "To" + " " + rs.Fields(2)
      m = m + 1 'problem so remove
          'starting if 3rd level
          rs1.MoveFirst
          Do While (Not rs1.EOF)
          
           If 0 <> InStr(Left(node2(z).Text, (InStr(node2(z).Text, "To") - 2)), rs1.Fields(1)) And 0 <> InStr(Right(node2(z).Text, (Len(node2(z).Text) - InStr(node2(z).Text, "To") - 2)), rs1.Fields(2)) Then
                   
             Set node3 = TreeView1.Nodes.Add(node2(z), tvwChild, , rs1.Fields(10))
              node3.Image = 4
                
          
          'starting of 4th level
                    
            If rs2.BOF = False Then
            rs2.MoveFirst
            End If
                    
                 Do While (Not rs2.EOF)
                  If (rs1.Fields(10) = rs2.Fields(14)) Then
              
                    Set node4 = TreeView1.Nodes.Add(node3, tvwChild, , rs2.Fields(13))
                     node4.Image = 6
                   End If
                           
                   rs2.MoveNext
                 Loop
          
          
          
          
          
           End If
          rs1.MoveNext
          Loop
          
     End If
     l = 0
     rs.MoveNext
Next j
Next z



'for expend the tree
TreeView1.Nodes(1).Expanded = True
Exit Function
err_h:
MsgBox " error while connecting to database"


End Function





Private Sub Image5_Click()
Frame1.Visible = False
Frame2.Visible = False
End Sub

Private Sub Image6_Click()
Frame1.Visible = False
Frame2.Visible = False
End Sub

Private Sub search_Click()
On Error GoTo err_h:

TreeView1.Refresh
TreeView1.Nodes(1).Expanded = False

Dim v As Integer, ul As Integer
v = 0
ul = 0
If Option1.Value = True Then
  For v = 1 To TreeView1.Nodes.Count
  If (Text1.Text = TreeView1.Nodes.Item(v)) Then
  'MsgBox "found"
  TreeView1.Nodes(1).Expanded = True
  'TreeView1.Nodes(1).Child.Expanded = True
  'TreeView1.Nodes(v).Expanded = True
  'TreeView1.Nodes.Item(v).Expanded = True
  'TreeView1.Nodes.Item(v).FirstSibling.FirstSibling.FirstSibling.Expanded = True
  TreeView1.Nodes.Item(v).FirstSibling.Parent.Expanded = True
  TreeView1.Nodes.Item(v).FirstSibling.Parent.FirstSibling.Parent.Expanded = True
  TreeView1.Nodes.Item(v).BackColor = RGB(9, 115, 103)
  TreeView1.Nodes.Item(v).ForeColor = RGB(250, 250, 250)
  ul = 1
 
  Exit Sub
  End If
  Next v

If (ul = 0) Then Image7.Visible = True
'MsgBox "harish"

End If
'************************************8

If Option2.Value = True Then
 For v = 1 To TreeView1.Nodes.Count
  If (Text1.Text = TreeView1.Nodes.Item(v)) Then
  'MsgBox "found"
  TreeView1.Nodes(1).Expanded = True
  'TreeView1.Nodes(1).Child.Expanded = True
  'TreeView1.Nodes(v).Expanded = True
  'TreeView1.Nodes.Item(v).Expanded = True
  'TreeView1.Nodes.Item(v).FirstSibling.FirstSibling.FirstSibling.Expanded = True
  TreeView1.Nodes.Item(v).FirstSibling.Parent.Expanded = True
  TreeView1.Nodes.Item(v).FirstSibling.Parent.FirstSibling.Parent.Expanded = True
  TreeView1.Nodes.Item(v).FirstSibling.Parent.FirstSibling.Parent.FirstSibling.Parent.Expanded = True
  TreeView1.Nodes.Item(v).BackColor = RGB(9, 115, 103)
  TreeView1.Nodes.Item(v).ForeColor = RGB(250, 250, 250)
  ul = 1
  Exit Sub
  End If
  Next v
If (ul = 0) Then Image7.Visible = True


'MsgBox "harish"
End If



Exit Sub
err_h:
MsgBox "error occureed while connecting to the database", vbInformation

End Sub

Private Sub Text1_Click()
Image7.Visible = False
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
'Text1.Text = Node.Text
'Text2.Text = Node.FullPath


'If TreeView1.Nodes(1).Child.Expanded = True Then
'MsgBox "SDfasf"
'End If

If Node.Image = 4 Then
Frame1.Visible = True
Frame2.Visible = False
rs.MoveFirst

Do While (Not rs.EOF)
If rs.Fields(10) = Node.Text Then

 Label2.Caption = rs.Fields(0)
 Label4.Caption = rs.Fields(1)
 Label6.Caption = rs.Fields(2)
 Label8.Caption = rs.Fields(3)
 Label10.Caption = rs.Fields(4)
 Label11.Caption = rs.Fields(5)
 Label13.Caption = rs.Fields(6)
 Label15.Caption = rs.Fields(7)
 Label17.Caption = rs.Fields(8)
 Label19.Caption = rs.Fields(9)
 Label21.Caption = rs.Fields(10)
 Exit Sub
 End If
rs.MoveNext

Loop
End If



'**********************************
Label26.Caption = ""
Label27.Caption = ""
Label28.Caption = ""
Label34.Caption = ""



If Node.Image = 6 Then
Frame1.Visible = False
Frame2.Visible = True
rs2.MoveFirst
Do While (Not rs2.EOF)
If rs2.Fields(13) = Node.Text Then

Label22.Caption = rs2.Fields(0)

Label23.Caption = rs2.Fields(1)
Label24.Caption = rs2.Fields(2)

Label26.Caption = rs2.Fields(3)
Label27.Caption = rs2.Fields(4)
Label28.Caption = rs2.Fields(5)




Label30.Caption = rs2.Fields(6)
Label32.Caption = rs2.Fields(7)


Label34.Caption = rs2.Fields(8)

Label36.Caption = rs2.Fields(9)
Label38.Caption = rs2.Fields(10)
Label40.Caption = rs2.Fields(11)
Label42.Caption = rs2.Fields(12)
Label44.Caption = rs2.Fields(13)
Label46.Caption = rs2.Fields(14)



 
 Exit Sub
 End If
rs2.MoveNext

Loop




End If



End Sub




