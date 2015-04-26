VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form6 
   BackColor       =   &H80000005&
   Caption         =   "Flight Entry Form"
   ClientHeight    =   9360
   ClientLeft      =   2205
   ClientTop       =   1320
   ClientWidth     =   9195
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   9195
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9120
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H80000005&
      Caption         =   "Backup and Restore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   1215
      Left            =   4680
      TabIndex        =   59
      Top             =   8040
      Width           =   4335
      Begin VB.CommandButton restore_backup 
         BackColor       =   &H80000005&
         Height          =   855
         Left            =   2400
         Picture         =   "flight entry form.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton create_backup 
         BackColor       =   &H80000005&
         Height          =   855
         Left            =   1080
         Picture         =   "flight entry form.frx":1B42
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H80000005&
      Caption         =   "Database Tree View"
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   120
      TabIndex        =   53
      Top             =   7680
      Width           =   4335
      Begin VB.CheckBox Check6 
         BackColor       =   &H80000005&
         Caption         =   "View Database in tree form"
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   360
         TabIndex        =   55
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H80000005&
         Height          =   1215
         Left            =   3000
         Picture         =   "flight entry form.frx":3684
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   120
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   5640
      Top             =   9360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "flight entry form.frx":66C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "flight entry form.frx":7318
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "flight entry form.frx":92EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "flight entry form.frx":A53D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "flight entry form.frx":B645
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "flight entry form.frx":10197
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "flight entry form.frx":108A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "flight entry form.frx":11EFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "flight entry form.frx":125EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "flight entry form.frx":12CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "flight entry form.frx":133A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "flight entry form.frx":14EF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "flight entry form.frx":17F44
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "flight entry form.frx":19A96
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H80000005&
      Caption         =   "Setting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   4680
      TabIndex        =   41
      Top             =   6480
      Width           =   4335
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "flight entry form.frx":1B5E8
         Left            =   960
         List            =   "flight entry form.frx":1B5F2
         TabIndex        =   62
         Text            =   "ORCL"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox driver 
         Height          =   315
         ItemData        =   "flight entry form.frx":1B600
         Left            =   2400
         List            =   "flight entry form.frx":1B60D
         TabIndex        =   57
         Text            =   "Oracle in XE"
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H80000005&
         Caption         =   "Create database (airlines_data) and Tables"
         Height          =   375
         Left            =   1320
         TabIndex        =   52
         Top             =   600
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CommandButton Create_database 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Create"
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton advance_setting 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Advance"
         Height          =   375
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton submit 
         Height          =   375
         Left            =   2880
         Picture         =   "flight entry form.frx":1B657
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2640
         TabIndex        =   48
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2640
         TabIndex        =   47
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton back 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   120
         Picture         =   "flight entry form.frx":1D619
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton database_setting 
         BackColor       =   &H80000005&
         Height          =   1095
         Left            =   2400
         Picture         =   "flight entry form.frx":1EC63
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton admin_setting 
         BackColor       =   &H80000005&
         Height          =   1095
         Left            =   1080
         Picture         =   "flight entry form.frx":1F34F
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000005&
         Caption         =   "Data Source"
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
         Left            =   1200
         TabIndex        =   58
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   120
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000005&
         Caption         =   "New Password"
         Height          =   375
         Left            =   1320
         TabIndex        =   46
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000005&
         Caption         =   "New ID"
         Height          =   375
         Left            =   1320
         TabIndex        =   45
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      Caption         =   "Customer Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1455
      Left            =   4680
      TabIndex        =   36
      Top             =   4920
      Width           =   4335
      Begin VB.CheckBox Check4 
         BackColor       =   &H80000005&
         Caption         =   "Create customer data report"
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Height          =   1215
         Left            =   3000
         Picture         =   "flight entry form.frx":1FA2E
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Flight Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1455
      Left            =   4680
      TabIndex        =   35
      Top             =   3360
      Width           =   4335
      Begin VB.CheckBox Check3 
         BackColor       =   &H80000005&
         Caption         =   "Create flight data report"
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   600
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Height          =   1215
         Left            =   3000
         Picture         =   "flight entry form.frx":2012C
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000005&
      Caption         =   "New Flight Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   7335
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Log Out"
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton clear_button 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clear"
         Height          =   495
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   975
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   2895
         Left            =   120
         TabIndex        =   4
         Top             =   3240
         Width           =   3615
         _Version        =   524288
         _ExtentX        =   6376
         _ExtentY        =   5106
         _StockProps     =   1
         BackColor       =   -2147483643
         Year            =   2012
         Month           =   7
         Day             =   25
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   7
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
      Begin VB.CommandButton submit_button 
         BackColor       =   &H00800000&
         Height          =   495
         Left            =   360
         Picture         =   "flight entry form.frx":24C6E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6600
         Width           =   1095
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H80000005&
         Caption         =   "Bussiness Class"
         Height          =   1335
         Left            =   120
         TabIndex        =   30
         Top             =   5040
         Width           =   4095
         Begin VB.ComboBox seat_business 
            Height          =   315
            ItemData        =   "flight entry form.frx":26C30
            Left            =   240
            List            =   "flight entry form.frx":26C52
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox price_business 
            Height          =   375
            Left            =   2520
            MaxLength       =   10
            TabIndex        =   10
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label16 
            BackColor       =   &H80000005&
            Caption         =   "No Of Seats :"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label15 
            BackColor       =   &H80000005&
            Caption         =   "Price Per Seat :"
            Height          =   255
            Left            =   2520
            TabIndex        =   31
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000005&
         Caption         =   "Economy Class"
         Height          =   1335
         Left            =   120
         TabIndex        =   27
         Top             =   3600
         Width           =   4095
         Begin VB.ComboBox seat_economy 
            Height          =   315
            ItemData        =   "flight entry form.frx":26C75
            Left            =   240
            List            =   "flight entry form.frx":26C97
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox price_economy 
            Height          =   375
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   8
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label14 
            BackColor       =   &H80000005&
            Caption         =   "No Of Seats :"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label13 
            BackColor       =   &H80000005&
            Caption         =   "Price Per Seat :"
            Height          =   255
            Left            =   2400
            TabIndex        =   28
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.ComboBox hour_box 
         Height          =   315
         ItemData        =   "flight entry form.frx":26CBA
         Left            =   2400
         List            =   "flight entry form.frx":26D06
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3000
         Width           =   495
      End
      Begin VB.ComboBox minutes_box 
         Height          =   315
         ItemData        =   "flight entry form.frx":26D60
         Left            =   3000
         List            =   "flight entry form.frx":26D73
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3000
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         DisabledPicture =   "flight entry form.frx":26D8A
         Height          =   375
         Left            =   1320
         Picture         =   "flight entry form.frx":272CC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox date_box 
         DragMode        =   1  'Automatic
         Height          =   405
         Left            =   120
         TabIndex        =   25
         Text            =   "mm/dd/yyyy"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.ComboBox to_box 
         Height          =   315
         ItemData        =   "flight entry form.frx":27F0E
         Left            =   2280
         List            =   "flight entry form.frx":27F2A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1680
         Width           =   1935
      End
      Begin VB.ComboBox from_box 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "flight entry form.frx":27F74
         Left            =   120
         List            =   "flight entry form.frx":27F90
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "harish"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.ComboBox select_flight 
         Height          =   315
         ItemData        =   "flight entry form.frx":27FDA
         Left            =   120
         List            =   "flight entry form.frx":27FED
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   720
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   4320
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   2880
         TabIndex        =   34
         Top             =   6720
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000005&
         Caption         =   "Flight No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   1800
         TabIndex        =   33
         Top             =   6720
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "Hours:   Minutes:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   26
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "Departure Time"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         Caption         =   "Departure Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000005&
         Caption         =   "To"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2280
         TabIndex        =   22
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000005&
         Caption         =   "From"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "Select Flight"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000005&
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   4680
      TabIndex        =   14
      Top             =   1560
      Width           =   4335
      Begin VB.CommandButton customer_delete 
         BackColor       =   &H80000005&
         Height          =   1335
         Left            =   3000
         Picture         =   "flight entry form.frx":28026
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H80000005&
         Caption         =   "Delete all old customer data"
         ForeColor       =   &H00FF8080&
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000005&
      Caption         =   "Flight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   4680
      TabIndex        =   13
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton flight_delete 
         BackColor       =   &H80000005&
         Height          =   1335
         Left            =   3000
         Picture         =   "flight entry form.frx":2911E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000005&
         Caption         =   "Delete all old flight data"
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   4560
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line4 
      X1              =   4560
      X2              =   9240
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line3 
      X1              =   4560
      X2              =   9120
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line2 
      X1              =   4560
      X2              =   4560
      Y1              =   0
      Y2              =   9120
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Const REG_SZ = 1    'Constant for a string variable type.
    Private Const HKEY_LOCAL_MACHINE = &H80000002

    Private Declare Function RegCreateKey Lib "advapi32.dll" Alias _
       "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
       phkResult As Long) As Long

    Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
       "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
       ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal _
       cbData As Long) As Long

    Private Declare Function RegCloseKey Lib "advapi32.dll" _
       (ByVal hKey As Long) As Long
'databse creator

Dim setting As Integer
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim flight_no As String
Dim val As Integer
Private Sub admin_setting_Click()
admin_setting.Visible = False
database_setting.Visible = False
back.Visible = True
Label8.Visible = True
Label9.Visible = True
Text1.Visible = True
Text2.Visible = True
submit.Visible = True
setting = 0
Text1.SetFocus
Text2.PasswordChar = "*"
End Sub

Private Sub advance_setting_Click()
Label8.Visible = False
Label9.Visible = False
Label10.Visible = True
Text1.Visible = False
Text2.Visible = False
advance_setting.Visible = False
submit.Visible = False
Create_database.Visible = True
Check5.Visible = True
driver.Visible = True
Combo1.Visible = True
End Sub

Private Sub back_Click()
admin_setting.Visible = True
database_setting.Visible = True

back.Visible = False
Label8.Visible = False
Label9.Visible = False
Text1.Visible = False
Text2.Visible = False
submit.Visible = False
Image1.Visible = False
advance_setting.Visible = False
Create_database.Visible = False
Check5.Visible = False
driver.Visible = False
Label10.Visible = False
Combo1.Visible = False
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
date_box.Text = ""
date_box.Text = Calendar1.Value
Calendar1.Visible = False
Else
date_box.Text = "mm/dd/yyyy"
End If
End Sub



Private Sub clear_button_Click()
date_box = "mm/dd/yyyy"
price_economy.Text = ""
price_business.Text = ""
Label12.Caption = ""


End Sub

Private Sub Command1_Click()
If Calendar1.Visible = True Then
  Calendar1.Visible = False
Else


Calendar1.Visible = True
End If

End Sub



Private Sub Command2_Click()
On Error GoTo err_h:
db.Close
Unload Me
Exit Sub

err_h:
Unload Me
End Sub


Private Sub Command3_Click()
If Check3.Value = 1 Then

flight_data_report.Show
Else
MsgBox "Please tick the check box", vbExclamation
End If
End Sub

Private Sub Command4_Click()
If Check4.Value = 1 Then
cust_data_report.Show
Else
MsgBox "Please tick the check box", vbExclamation
End If
End Sub



Private Sub Command5_Click()
On Error Resume Next:
If Check6.Value = 1 Then
Form10.Visible = True
Else

MsgBox "Please tick the check box", vbExclamation

End If

End Sub

Private Sub Command6_Click()

End Sub

Private Sub create_backup_Click()
Dim day As Variant, month As Variant, year As Variant, add As Variant
On Error Resume Next

CommonDialog1.Filter = "Backup file (*.backup)|*.backup"
CommonDialog1.FileName = Date$
CommonDialog1.DialogTitle = "select the location for Backup file"
CommonDialog1.Flags = 0
CommonDialog1.ShowSave
Kill CommonDialog1.FileName 'check for existing files fo backup

 If CommonDialog1.Flags <> 0 Then
  rs.Open "select * from flight_data", db, adOpenDynamic, adLockOptimistic, adCmdText
  rs1.Open "select * from cust_data", db, adOpenDynamic, adLockOptimistic, adCmdText
 
  rs.MoveFirst

  Open CommonDialog1.FileName For Append As #1
 
  Do While (Not rs.EOF)
  Write #1, rs.Fields(0), rs.Fields(1), rs.Fields(2), rs.Fields(3), rs.Fields(4), rs.Fields(5), rs.Fields(6), rs.Fields(7), rs.Fields(8), rs.Fields(9), rs.Fields(10)
  rs.MoveNext
 
      If (rs.EOF) Then
      Write #1, "***", "***", "***", "***", "***", "***", "***", "***", "***", "***", "***"
      rs1.MoveFirst
      Do While (Not rs1.EOF)
      If IsNull(rs1.Fields(3)) Then day = "blank"
      If IsNull(rs1.Fields(4)) Then month = "blank"
      If IsNull(rs1.Fields(5)) Then year = "blank"
      If IsNull(rs1.Fields(8)) Then add = "blank"
       Write #1, rs1.Fields(0), rs1.Fields(1), rs1.Fields(2), day, month, year, rs1.Fields(6), rs1.Fields(7), add, rs1.Fields(9), rs1.Fields(10), rs1.Fields(11), rs1.Fields(12), rs1.Fields(13), rs1.Fields(14), rs1.Fields(15)
      
      
      rs1.MoveNext
      Loop
    End If
 
 
  Loop
  MsgBox "Backup creation sucessful", vbInformation
  Close #1
  rs.Close
  rs1.Close
Else
Exit Sub
End If
End Sub

Private Sub Create_database_Click()

Dim db As New ADODB.Connection
If Check5.Value = 1 Then

  On Error GoTo err_h:




'Button

'***********

 Dim DataSourceName As String
   Dim DatabaseName As String
   Dim Description As String
   Dim DriverPath As String
   Dim DriverName As String
   Dim LastUser As String
   Dim Regional As String
   Dim Server As String

   Dim lResult As Long
   Dim hKeyHandle As Long

   'Specify the DSN parameters.

   DataSourceName = "airlines_data"
   DatabaseName = "airlines_data"
   Description = "airlines table"
   DriverPath = ""
   LastUser = ""
   Server = Combo1.Text
   DriverName = driver.Text

   'Create the new DSN key.

   lResult = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & _
        DataSourceName, hKeyHandle)

   'Set the values of the new DSN key.

   lResult = RegSetValueEx(hKeyHandle, "Database", 0&, REG_SZ, _
      ByVal DatabaseName, Len(DatabaseName))
   lResult = RegSetValueEx(hKeyHandle, "Description", 0&, REG_SZ, _
      ByVal Description, Len(Description))
   lResult = RegSetValueEx(hKeyHandle, "Driver", 0&, REG_SZ, _
      ByVal DriverPath, Len(DriverPath))
   lResult = RegSetValueEx(hKeyHandle, "LastUser", 0&, REG_SZ, _
      ByVal LastUser, Len(LastUser))
   lResult = RegSetValueEx(hKeyHandle, "Server", 0&, REG_SZ, _
      ByVal Server, Len(Server))

   'Close the new DSN key.

   lResult = RegCloseKey(hKeyHandle)

   'Open ODBC Data Sources key to list the new DSN in the ODBC Manager.
   'Specify the new value.
   'Close the key.

   lResult = RegCreateKey(HKEY_LOCAL_MACHINE, _
      "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKeyHandle)
   lResult = RegSetValueEx(hKeyHandle, DataSourceName, 0&, REG_SZ, _
      ByVal DriverName, Len(DriverName))
   lResult = RegCloseKey(hKeyHandle)
   MsgBox "Database Creation sucessful"
   Call create_table
   MsgBox "Now change the uid and password of database connectivity"
back_Click
database_setting_Click 'fsaf'
Exit Sub
err_h:
MsgBox "ERROR occured while connecting to oracle database or oracle not installed"

Else

MsgBox "Tick the check box for creating database", vbInformation

End If

End Sub

Private Sub customer_delete_Click()
On Error GoTo err_h:

Dim tempi As Integer
Dim counter As Integer
counter = 0
tempi = 0
If Check2.Value = 1 Then
 
 rs1.Open "select * from cust_data ", db, adOpenDynamic, adLockOptimistic, adCmdText
If (Not rs1.BOF) Then
rs1.MoveFirst
Else
MsgBox "Customer DataBase is empty", vbInformation
rs1.Close
Exit Sub
End If


rs.Open "select * from flight_data", db, adOpenDynamic, adLockOptimistic, adCmdText
If (rs.EOF) Then

   Do While (Not rs1.EOF)
   rs1.Delete
   rs1.MoveNext
   counter = counter + 1
   Loop
MsgBox counter & " customer data deleted"
rs.Close
rs1.Close

Exit Sub
Else
rs.MoveFirst
End If





 
Do While (Not rs1.EOF)

   Do While (Not rs.EOF)
   
    If 0 = StrComp(rs("flight_no").Value, rs1("flight_no").Value) Then
     tempi = 1
    rs.MoveNext
    Else
    
    rs.MoveNext
    End If
   
    If rs.EOF And tempi = 0 Then
      counter = counter + 1
      rs1.Delete
       tempi = 0
       
      
    End If
    Loop

rs.MoveFirst
rs1.MoveNext
Loop

MsgBox counter & " customer data deleted", vbInformation
rs.Close
rs1.Close


Else
MsgBox "Please tick the check box", vbExclamation
End If
Exit Sub
err_h:

MsgBox "error occureed while connecting to the database", vbInformation

End Sub

Private Sub database_setting_Click()
admin_setting.Visible = False
database_setting.Visible = False
back.Visible = True
Label8.Visible = True
Label9.Visible = True
Text1.Visible = True
Text2.Visible = True
submit.Visible = True
Image1.Visible = True
advance_setting.Visible = True
Text1.SetFocus
setting = 1
Text2.PasswordChar = ""
End Sub

Private Sub flight_delete_Click()
On Error GoTo err_h:

Dim tempdate As Date
Dim tempi As Integer


tempi = 0
tempdate = Date
tempdate = DateAdd("d", -1, tempdate)

If Check1.Value = 1 Then ' means checked
  rs.Open "select * from flight_data", db, adOpenDynamic, adLockOptimistic, adCmdText
  If (Not rs.BOF) Then
    rs.MoveFirst
  Else
  MsgBox "Flight DataBase is empty", vbInformation
  rs.Close
  Exit Sub
  End If
  
  
  Do While (Not rs.EOF)
  If rs("depart_date").Value < tempdate Then
  rs.Delete
  tempi = tempi + 1
  End If
  rs.MoveNext
  Loop
  'rs.Close

  MsgBox tempi & " flight deleted", vbInformation
rs.Close
Else
MsgBox "Please Tick the Check box", vbExclamation

End If
Exit Sub
err_h:

MsgBox "error occureed while connecting to the database", vbInformation


End Sub

Private Sub Form_Click()

Calendar1.Visible = False
End Sub

Private Sub Form_Load()
Calendar1.Value = Date
Command1.Picture = ImageList.ListImages(1).Picture
submit_button.Picture = ImageList.ListImages(2).Picture
flight_delete.Picture = ImageList.ListImages(3).Picture
customer_delete.Picture = ImageList.ListImages(4).Picture
Command3.Picture = ImageList.ListImages(5).Picture
Command4.Picture = ImageList.ListImages(6).Picture
submit.Picture = ImageList.ListImages(2).Picture
back.Picture = ImageList.ListImages(7).Picture
admin_setting.Picture = ImageList.ListImages(8).Picture
database_setting.Picture = ImageList.ListImages(10).Picture
Image1.Picture = ImageList.ListImages(11).Picture
Command5.Picture = ImageList.ListImages(12).Picture
create_backup.Picture = ImageList.ListImages(13).Picture
restore_backup.Picture = ImageList.ListImages(14).Picture
'db.ConnectionString = "dsn=airlines_data;uid=system;pwd=harish"

On Error GoTo err_h:
db.ConnectionString = "dsn=airlines_data;uid=" & user_name & ";pwd=" & pass_word & " "
db.Open
Calendar1.Visible = False
Exit Sub
err_h:

Calendar1.Visible = False
MsgBox "u have to need change database password and user name or fix the connectivity problem", vbInformation
Timer1.Enabled = True
End Sub

Public Function check_all_are_fill()
val = 0
If select_flight.Text = "" Then
val = 1
End If

If from_box.Text = "" Then
val = 1
End If

If to_box.Text = "" Then
val = 1
End If

If date_box.Text = "mm/dd/yyyy" Then
val = 1
End If

If hour_box.Text = "" Or minutes_box.Text = "" Then
val = 1
End If

If seat_economy.Text = "" Or price_economy.Text = "" Or seat_business.Text = "" Or price_business.Text = "" Then
val = 1
End If


If from_box.Text = to_box.Text Then
MsgBox "going to same city not possible"
val = 1
End If
End Function





Private Sub price_business_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyBack Then
Exit Sub
End If
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
End Sub

Private Sub price_economy_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyBack Then
Exit Sub
End If
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
End Sub

Private Sub restore_backup_Click()
Dim a As String, b As String, c As String, d As String, e As Variant, f As Variant, g As Variant, h As Variant, i As Variant, j As Variant, k As Variant
Dim aa As String, ab As String, ac As String, ad As Variant, ae As String, af As Variant, ag As Variant, ah As Variant, ai As Variant, aj As Variant, ak As Variant, al As Variant, am As Variant, an As Variant, ao As Variant, ap As Variant

CommonDialog1.Filter = "Backup file (*.backup)|*.backup"
CommonDialog1.DialogTitle = "Select the airlines data backup file"
CommonDialog1.Flags = 0
CommonDialog1.ShowOpen
'MsgBox CommonDialog1.FileName
'MsgBox CommonDialog1.Flags = 1
If CommonDialog1.Flags <> 0 Then
'MsgBox "opened"
'MsgBox "adF"
 On Error GoTo err_h:
 
 Open CommonDialog1.FileName For Input As #1
 db.Execute "delete from flight_data"

 Do While (1)
 Input #1, a, b, c, d, e, f, g, h, i, j, k
 
 If a = "***" Then
   db.Execute "delete from cust_data"
    Do While (1)
         
         Input #1, aa, ab, ac, ad, ae, af, ag, ah, ai, aj, ak, al, am, an, ao, ap
         ''If p Then MsgBox "har"
    
           If ad = "blank" Then ad = ""
           If ae = "blank" Then ae = ""
           If af = "blank" Then af = ""
           If ai = "blank" Then ai = ""
          
          rs1.Open "insert into cust_data values('" & aa & "','" & ab & "','" & ac & "','" & ad & "','" & ae & "','" & af & "','" & ag & "','" & ah & "','" & ai & "','" & aj & "','" & ak & "','" & al & "','" & am & "','" & an & "','" & ao & "','" & ap & "')", db, adOpenDynamic, adLockOptimistic, adCmdText
     Loop
    End If
 
  rs.Open "insert into flight_data values('" & a & "','" & b & "','" & c & "','" & d & "','" & e & "','" & f & "','" & g & "','" & h & "','" & i & "','" & j & "','" & k & "')", db, adOpenDynamic, adLockOptimistic, adCmdText
 Loop
 Close #1
'MsgBox "backup restore sucessful", vbInformation
Else
Exit Sub

End If

Exit Sub
err_h:
''rs.Close
MsgBox "backup restore sucessful", vbInformation
Close #1

End Sub

Private Sub submit_button_Click()
On Error GoTo err_h:
Call check_all_are_fill

If val = 1 Then
  MsgBox "please fill all the entry", vbExclamation
  Exit Sub
  
Else


flight_no = Left(select_flight.Text, 3) & Calendar1.day & Calendar1.month & Calendar1.year & CInt(Int(Rnd() * Int(Rnd() * 199)))
rs.Open "insert into flight_data values('" & select_flight.Text & "','" & from_box.Text & "','" & to_box.Text & "','" & date_box.Text & "','" & hour_box.Text & "','" & minutes_box.Text & "','" & seat_economy.Text & "','" & price_economy.Text & "','" & seat_business.Text & "','" & price_business.Text & "','" & flight_no & "')", db, adOpenDynamic, adLockOptimistic, adCmdText

Label12.Caption = flight_no
End If
Exit Sub
err_h:
MsgBox "error occureed while connecting to the database", vbInformation


End Sub


Private Sub submit_Click()
On Error GoTo err_h:
Dim valdata As Variant
Dim valdata1 As Variant

If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Fill both of box", vbInformation
Exit Sub
End If

If setting = 0 Then
''login code'
Call change_password

Else
Open "database_connectivity.dat" For Binary As #1
valdata = Text1.Text
valdata1 = Text2.Text
Put #1, , valdata
Put #1, , valdata1
Close #1
MsgBox "Changes Sucessful", vbInformation

Open "database_connectivity.dat" For Binary As #1
Get #1, , user_name
Get #1, , pass_word
Close #1
If db.State Then
db.Close
End If


db.ConnectionString = "dsn=airlines_data;uid=" & user_name & ";pwd=" & pass_word & " "

db.Open
End If
valdata = ""
valdata1 = ""

admin_setting.Visible = True
database_setting.Visible = True
back.Visible = False
Label8.Visible = False
Label9.Visible = False
Text1.Visible = False
Text2.Visible = False
submit.Visible = False
Text1.Text = ""
Text2.Text = ""
Timer1.Enabled = False
database_setting.Picture = ImageList.ListImages(10).Picture
Image1.Visible = False
advance_setting.Visible = False


Exit Sub
err_h:
MsgBox "error occureed while reconnecting to the database", vbInformation
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
If (database_setting.Picture = ImageList.ListImages(9).Picture) Then
database_setting.Picture = ImageList.ListImages(10).Picture
Else
database_setting.Picture = ImageList.ListImages(9).Picture
End If



End Sub

Public Function create_table()
Dim db As New ADODB.Connection
On Error GoTo erroh:
'db.ConnectionString = "dsn=airlines_data;uid=system;pwd=harish"
'db.Open
db.ConnectionString = "dsn=airlines_data;uid=" & user_name & ";pwd=" & pass_word & " "
db.Open
db.Execute "Create TABLE cust_data (title varchar(4),first_name varchar(20),last_name varchar(20),date_of_birth_day number(2),date_of_birth_month char(4),date_of_birth_year number(4),mobile_no char(15),address_line_one varchar(50),address_line_two varchar(50),city varchar(10),state varchar(15),country varchar(15),pincode char(8),pnr_no varchar(20),flight_no varchar(15),cabin varchar(10))"
db.Execute "create table flight_data(flight_company varchar(15),depart_city varchar(10),arrival_city varchar(10),depart_date varchar(10),depart_hour number(3),depart_minute number(3),eco_no_of_seat number(4),eco_price varchar(10),busi_no_of_seat number(4),busi_price varchar(10),flight_no varchar(15) Primary Key)"
db.Execute "create table login(username varchar(10),password varchar(20))"
db.Execute "insert into login values ( 'harish','harish')"



db.Close

MsgBox "tables successfull create"
Exit Function
erroh:
MsgBox "error occured while creating tables or Tables already created"

End Function







Public Sub change_password()
On Error GoTo err_h:

If login_uname = "key" And login_passwd = "key" Then
MsgBox "login again"
Exit Sub
End If



Dim valdata As Variant
Dim valdata1 As Variant
valdata = Text1.Text
valdata1 = Text2.Text

db.Execute "delete from Login where username = '" & login_uname & "' and password= '" & login_passwd & "' "
db.Execute "insert into login values('" & valdata & "','" & valdata1 & "')"
login_uname = Text1.Text
login_passwd = Text2.Text
MsgBox "username and password sucessfully changed", vbInformation

Exit Sub
err_h:
MsgBox "User name and password not found", vbInformation
End Sub
