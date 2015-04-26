VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   Caption         =   "Airlines Reservation System"
   ClientHeight    =   8925
   ClientLeft      =   1425
   ClientTop       =   1410
   ClientWidth     =   10395
   DrawStyle       =   2  'Dot
   Icon            =   "Airlines Reservation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "Airlines Reservation.frx":0442
   ScaleHeight     =   8925
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   6000
      Top             =   5520
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   6120
      Top             =   7680
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2775
      Left            =   120
      TabIndex        =   25
      Top             =   6120
      Width           =   3855
      _Version        =   524288
      _ExtentX        =   6800
      _ExtentY        =   4895
      _StockProps     =   1
      BackColor       =   -2147483643
      Year            =   2012
      Month           =   8
      Day             =   28
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   0
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000005&
      Caption         =   "Database"
      Height          =   2175
      Left            =   240
      TabIndex        =   22
      Top             =   6600
      Visible         =   0   'False
      Width           =   6855
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1935
         Left            =   1560
         TabIndex        =   26
         Top             =   120
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   3413
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000005&
         Caption         =   "Customer Detail"
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000005&
         Caption         =   "Flight Detail"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command4 
      Height          =   255
      Left            =   120
      Picture         =   "Airlines Reservation.frx":068B
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6120
      Width           =   735
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000005&
      Caption         =   "Print Ticket"
      Height          =   2655
      Left            =   7320
      TabIndex        =   20
      Top             =   6000
      Width           =   2895
      Begin VB.CommandButton Command3 
         Height          =   495
         Left            =   960
         Picture         =   "Airlines Reservation.frx":08DB
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   840
         Width           =   1695
      End
      Begin VB.Image Image9 
         Height          =   390
         Left            =   1320
         Picture         =   "Airlines Reservation.frx":289D
         Top             =   360
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Image Image8 
         Height          =   450
         Left            =   960
         Picture         =   "Airlines Reservation.frx":43AF
         Top             =   360
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000005&
         Caption         =   "PNR NO"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      Caption         =   "cancellation"
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   7440
      TabIndex        =   18
      Top             =   3240
      Width           =   2775
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Height          =   495
         Left            =   840
         Picture         =   "Airlines Reservation.frx":6DA9
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Image Image3 
         Height          =   450
         Left            =   960
         Picture         =   "Airlines Reservation.frx":8D6B
         Top             =   360
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   1320
         Picture         =   "Airlines Reservation.frx":B765
         Top             =   360
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000005&
         Caption         =   "PNR NO"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Cabin/Class"
      Height          =   975
      Left            =   4320
      TabIndex        =   17
      Top             =   4080
      Width           =   2775
      Begin VB.OptionButton business_cabin 
         BackColor       =   &H80000005&
         Caption         =   "Business"
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton economy_cabin 
         BackColor       =   &H80000005&
         Caption         =   "Economy"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   10560
      Top             =   8160
   End
   Begin VB.PictureBox Picturebox 
      DrawStyle       =   1  'Dash
      Height          =   3015
      Left            =   0
      Picture         =   "Airlines Reservation.frx":D277
      ScaleHeight     =   2955
      ScaleWidth      =   7155
      TabIndex        =   15
      Top             =   0
      Width           =   7215
   End
   Begin VB.CommandButton show_flight_button 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Picture         =   "Airlines Reservation.frx":53FD1
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      DisabledPicture =   "Airlines Reservation.frx":55F93
      Height          =   375
      Left            =   1320
      Picture         =   "Airlines Reservation.frx":564D5
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   375
   End
   Begin VB.ComboBox to_box 
      Height          =   315
      ItemData        =   "Airlines Reservation.frx":57117
      Left            =   2160
      List            =   "Airlines Reservation.frx":57133
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4560
      Width           =   1935
   End
   Begin VB.ComboBox from_box 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "Airlines Reservation.frx":5717D
      Left            =   0
      List            =   "Airlines Reservation.frx":57199
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "harish"
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox text1 
      DragMode        =   1  'Automatic
      Height          =   405
      Left            =   120
      TabIndex        =   16
      Text            =   "mm/dd/yyyy"
      Top             =   5640
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   512
      ImageHeight     =   208
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Airlines Reservation.frx":571E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Airlines Reservation.frx":A5235
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Airlines Reservation.frx":EBF9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Airlines Reservation.frx":132D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Airlines Reservation.frx":179A73
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Airlines Reservation.frx":1C7AC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Airlines Reservation.frx":215B17
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Airlines Reservation.frx":217639
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Airlines Reservation.frx":21A043
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Airlines Reservation.frx":21AC95
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Airlines Reservation.frx":21CC67
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Airlines Reservation.frx":21CFC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Airlines Reservation.frx":21D227
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Airlines Reservation.frx":21D92D
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Airlines Reservation.frx":21F8FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Airlines Reservation.frx":2218D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Airlines Reservation.frx":23D77B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image12 
      Height          =   450
      Left            =   480
      Top             =   3600
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Image Image11 
      Height          =   570
      Left            =   4920
      Picture         =   "Airlines Reservation.frx":24100D
      Top             =   5640
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image Image10 
      Height          =   570
      Left            =   3120
      Picture         =   "Airlines Reservation.frx":24135D
      Top             =   7320
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Line Line13 
      X1              =   7200
      X2              =   10320
      Y1              =   8760
      Y2              =   8760
   End
   Begin VB.Line Line4 
      X1              =   7200
      X2              =   7200
      Y1              =   5160
      Y2              =   8760
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   7200
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Image Image7 
      Height          =   390
      Left            =   960
      Picture         =   "Airlines Reservation.frx":2416AD
      Top             =   5160
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Image Image6 
      Height          =   390
      Left            =   5640
      Picture         =   "Airlines Reservation.frx":2431BF
      Top             =   3720
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Image Image5 
      Height          =   390
      Left            =   2880
      Picture         =   "Airlines Reservation.frx":244CD1
      Top             =   4080
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Image Image4 
      Height          =   390
      Left            =   960
      Picture         =   "Airlines Reservation.frx":2467E3
      Top             =   4080
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Line Line12 
      X1              =   7440
      X2              =   7440
      Y1              =   0
      Y2              =   3120
   End
   Begin VB.Line Line11 
      X1              =   7320
      X2              =   10320
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line10 
      X1              =   7320
      X2              =   7320
      Y1              =   3120
      Y2              =   5880
   End
   Begin VB.Line Line9 
      X1              =   4200
      X2              =   4200
      Y1              =   3960
      Y2              =   5160
   End
   Begin VB.Image Image1 
      Height          =   3105
      Left            =   7560
      Picture         =   "Airlines Reservation.frx":2482F5
      Top             =   0
      Width           =   2760
   End
   Begin VB.Line Line8 
      X1              =   7200
      X2              =   7200
      Y1              =   3120
      Y2              =   6120
   End
   Begin VB.Line Line7 
      X1              =   10320
      X2              =   10320
      Y1              =   8760
      Y2              =   0
   End
   Begin VB.Line Line6 
      X1              =   7320
      X2              =   7320
      Y1              =   3120
      Y2              =   0
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   10200
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   7200
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7200
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000005&
      Caption         =   "Depart Date"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      Caption         =   "To"
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "From"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "BOOK YOUR DOMESTIC FLIGHT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   4215
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu index 
         Caption         =   "How to"
         Begin VB.Menu helpconnection 
            Caption         =   "Connect with database"
         End
         Begin VB.Menu helpbook 
            Caption         =   "Book a ticket"
         End
         Begin VB.Menu helpreprint 
            Caption         =   "Reprint the ticket"
         End
         Begin VB.Menu helpcancel 
            Caption         =   "Cancel a ticket"
         End
         Begin VB.Menu helpupdate 
            Caption         =   "Update the flight"
         End
         Begin VB.Menu helpdatabase 
            Caption         =   "Delete the old database"
         End
         Begin VB.Menu helpbackup 
            Caption         =   "Create and Restore backup"
         End
         Begin VB.Menu helpchange 
            Caption         =   "change login uid and password"
         End
         Begin VB.Menu helptreeview 
            Caption         =   "See the database tree view"
         End
         Begin VB.Menu helpmaster 
            Caption         =   "what is master password"
         End
      End
      Begin VB.Menu helphelp 
         Caption         =   "Help"
      End
      Begin VB.Menu about 
         Caption         =   "&About"
         Shortcut        =   ^O
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strSQL As String
Dim i As Integer
Dim time As Integer






Private Sub about_Click()
frmAbout.Show
End Sub

Private Sub business_cabin_Click()
Image6.Visible = False
cabin = "business"
End Sub

Private Sub Calendar1_Click()
Dim temp As Integer
temp = 0



If Calendar1.year < DatePart("yyyy", Now) Then
MsgBox "Wrong Date", vbInformation

     temp = 1
     
Else

      If Calendar1.month < DatePart("m", Now) And Calendar1.year <= DatePart("yyyy", Now) Then
         
         MsgBox "Wrong Date", vbInformation
        
          temp = 1
          Exit Sub
      End If
      
          If Calendar1.day < DatePart("d", Now) And Calendar1.month <= DatePart("m", Now) And Calendar1.year <= DatePart("yyyy", Now) Then
           MsgBox "Wrong Date", vbInformation
          
           temp = 1
         End If
End If



If temp = 0 Then
Text1.Text = ""
Text1.Text = Calendar1.Value
Calendar1.Visible = False
Else
Text1.Text = "mm/dd/yyyy"
End If

End Sub



Private Sub Command1_Click()

If Calendar1.Visible = True Then
   Calendar1.Visible = False
Else
Calendar1.Visible = True
End If
End Sub

Private Sub Command1_GotFocus()
Image7.Visible = False
End Sub

Private Sub Command2_Click()
'Form4.Visible = False
On Error GoTo err_h:
If Text2.Text = "" Then
  Image2.Visible = True
  Exit Sub
End If
  

pnr_cancel = Text2.Text

  'Form4.Visible = True
  Load Form4
  

If pnr_cancel = "notfound" Then
Image3.Visible = True
Unload Form4
Exit Sub
End If

Exit Sub
err_h:

MsgBox "error occureed while connecting to the database", vbExclamation


  
End Sub

Private Sub Command3_Click()
On Error GoTo err_h:
If Text3.Text = "" Then
Image9.Visible = True
Exit Sub
End If



Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset


db.ConnectionString = "dsn=airlines_data;uid=" & user_name & ";pwd=" & pass_word & " "
db.Open
rs.Open "select * from cust_data where pnr_no =  '" & Text3.Text & "'    ", db, adOpenDynamic, adLockOptimistic, adCmdText

If rs.EOF Then
Image8.Visible = True
Else
rs1.Open "select * from flight_data where flight_no =  '" & rs.Fields("flight_no") & "'    ", db, adOpenDynamic, adLockOptimistic, adCmdText

going_from = rs1.Fields("depart_city")

 going_to = rs1.Fields("arrival_city")
depart_date = rs1.Fields("depart_Date")
 cabin = rs.Fields("cabin")
flight_no = rs.Fields("flight_no")
first_name = rs.Fields("first_name")
last_name = rs.Fields("last_name")
 pnr_no = rs.Fields("pnr_no")
 If (0 = StrComp(cabin, "economy")) Then
   price = rs1.Fields("eco_price")
   Else
   price = rs1.Fields("busi_price")
   End If
   
depart_hour = rs1.Fields("depart_hour")
depart_minute = rs1.Fields("depart_minute")
rs1.Close
rs.Close
db.Close
dup_ticket = "true"
Load Form7
 Form7.Visible = True

End If
Exit Sub
err_h:

MsgBox "error occureed while connecting to the database", vbExclamation





End Sub

Private Sub Command4_Click()
If Frame4.Visible = True Then
Frame4.Visible = False
Command4.Picture = ImageList.ListImages(12).Picture
Exit Sub
End If
Image10.Visible = True
Option1.Value = True
Command4.Picture = ImageList.ListImages(13).Picture
Timer2.Enabled = True
'Frame4.Visible = True
Option2.Value = True


End Sub



Private Sub Command5_Click()
Load Form9



End Sub

Private Sub economy_cabin_Click()
Image6.Visible = False
cabin = "economy"
End Sub



Private Sub exit_Click()
Unload Form1
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload Form6
Unload Form7
Unload Form8
Unload frmAbout
Unload frmSplash

End Sub

Private Sub Form_Click()



Calendar1.Visible = False
End Sub
Private Sub Form_Load()
'Unload frmSplash
Image2.Picture = ImageList.ListImages(7).Picture
Image4.Picture = ImageList.ListImages(7).Picture
Image5.Picture = ImageList.ListImages(7).Picture
Image6.Picture = ImageList.ListImages(7).Picture
Image7.Picture = ImageList.ListImages(7).Picture
Image9.Picture = ImageList.ListImages(7).Picture
Image3.Picture = ImageList.ListImages(8).Picture
Image8.Picture = ImageList.ListImages(8).Picture
Command1.Picture = ImageList.ListImages(9).Picture
show_flight_button.Picture = ImageList.ListImages(10).Picture
Image10.Picture = ImageList.ListImages(11).Picture
Image11.Picture = ImageList.ListImages(11).Picture
Command4.Picture = ImageList.ListImages(12).Picture
Command2.Picture = ImageList.ListImages(14).Picture
Command3.Picture = ImageList.ListImages(15).Picture
Image1.Picture = ImageList.ListImages(16).Picture
Image12.Picture = ImageList.ListImages(17).Picture

Calendar1.Visible = False
time = 10
Picturebox.Picture = ImageList.ListImages(2).Picture



Open "database_connectivity.dat" For Binary As #1
Get #1, , user_name
Get #1, , pass_word

Close #1


Calendar1.Value = Date

'App.HelpFile = "C:\Documents and Settings\Administrator\Desktop\airlines reservation system project sql new\HELP1.HLP"
'Shell "winhlp32.exe -i connect help1.hlp"
End Sub







Private Sub from_box_GotFocus()
Image4.Visible = False
Image12.Visible = False
End Sub

Private Sub helpbackup_Click()
Shell "winhlp32.exe -i backup help1.hlp"
End Sub

Private Sub helpbook_Click()
Shell "winhlp32.exe -i booking help1.hlp"
End Sub

Private Sub helpcancel_Click()
Shell "winhlp32.exe -i cancel help1.hlp"
End Sub

Private Sub helpchange_Click()
Shell "winhlp32.exe -i uidpassword help1.hlp"
End Sub

Private Sub helpconnection_Click()
Shell "winhlp32.exe -i connect help1.hlp"
End Sub

Private Sub helpdatabase_Click()
Shell "winhlp32.exe -i deletion help1.hlp"
End Sub

Private Sub helphelp_Click()
Shell "winhlp32.exe  help1.hlp"
End Sub

Private Sub helpmaster_Click()
Shell "winhlp32.exe -i masterpassword help1.hlp"
End Sub

Private Sub helpreprint_Click()
Shell "winhlp32.exe -i reprint help1.hlp"
End Sub

Private Sub helptreeview_Click()
Shell "winhlp32.exe -i treeview help1.hlp"
End Sub

Private Sub helpupdate_Click()
Shell "winhlp32.exe -i update help1.hlp"
End Sub

Private Sub Image1_Click()
'Form5.Visible = True
Form6.Visible = True
End Sub





Private Sub Option1_GotFocus()
strSQL = "SELECT * FROM flight_Data"
Call show_database
End Sub
Private Sub Option2_Click()
strSQL = "SELECT * FROM cust_data"
Call show_database
End Sub


Private Sub show_flight_button_Click()
i = 0
Call check_all_fill

If i = 0 Then
going_from = from_box.Text
going_to = to_box.Text
depart_date = Text1.Text
Image11.Visible = True
Timer3.Enabled = True

Exit Sub
End If

End Sub
Private Sub Text2_Click()
Image2.Visible = False
Image3.Visible = False
End Sub



Private Sub Text3_Click()
Image8.Visible = False
Image9.Visible = False
End Sub



Private Sub Timer1_Timer()
Call imagechange
End Sub

Public Function imagechange()
time = time + 10

If time = 10 Then
Picturebox.Picture = ImageList.ListImages(1).Picture
End If

If time = 20 Then
Picturebox.Picture = ImageList.ListImages(2).Picture
End If

If time = 30 Then
Picturebox.Picture = ImageList.ListImages(3).Picture
End If

If time = 40 Then
Picturebox.Picture = ImageList.ListImages(4).Picture
End If

If time = 50 Then
Picturebox.Picture = ImageList.ListImages(5).Picture
End If

If time = 60 Then
Picturebox.Picture = ImageList.ListImages(6).Picture
time = 0
End If

End Function






Public Function check_all_fill()
  If from_box.Text = "" Then
  i = 1
  Image4.Visible = True
   End If

 If to_box.Text = "" Then
 i = 1
 Image5.Visible = True
 
 End If
 
 If Text1.Text = "mm/dd/yyyy" Then
  i = 1
  Image7.Visible = True
  End If

If economy_cabin.Value = False And business_cabin.Value = False Then
i = 1
Image6.Visible = True
End If

If from_box.Text = to_box.Text And from_box.Text <> "" Then
i = 1
Image12.Visible = True
End If



End Function



Private Sub Timer2_Timer()


Frame4.Visible = True
Image10.Visible = False
Timer2.Enabled = False

End Sub

Private Sub Timer3_Timer()
On Error GoTo err_h:
Form1.Visible = False
Form2.Visible = True
Image11.Visible = False
Timer3.Enabled = False
Exit Sub
err_h:

Form1.Visible = True
Image11.Visible = False
Timer3.Enabled = False
MsgBox "error occureed while connecting to the database", vbExclamation

End Sub

Private Sub to_box_GotFocus()
Image5.Visible = False
Image12.Visible = False
End Sub

Public Function show_database()
Dim oconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
On Error GoTo err_h:

oconn.ConnectionString = "dsn=airlines_data;uid=" & user_name & ";pwd=" & pass_word & " "

oconn.Open
rs.CursorType = adOpenStatic
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.Open strSQL, oconn, , , adCmdText
Set DataGrid1.DataSource = rs
DataGrid1.ReBind

Exit Function
err_h:
MsgBox "error is occured while connecting to database", vbExclamation



End Function
