VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BackColor       =   &H80000005&
   Caption         =   "Flight"
   ClientHeight    =   4290
   ClientLeft      =   2775
   ClientTop       =   2850
   ClientWidth     =   7740
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7740
   Begin MSComctlLib.ImageList ImageList 
      Left            =   8520
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   90
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flight info.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flight info.frx":0460
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   360
      Picture         =   "Flight info.frx":2822
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   5160
      Picture         =   "Flight info.frx":2C72
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Line Line17 
      X1              =   1680
      X2              =   1680
      Y1              =   120
      Y2              =   720
   End
   Begin VB.Line Line16 
      X1              =   240
      X2              =   1680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line15 
      X1              =   240
      X2              =   1680
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line14 
      X1              =   240
      X2              =   240
      Y1              =   120
      Y2              =   720
   End
   Begin VB.Line Line13 
      X1              =   5040
      X2              =   6840
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line12 
      X1              =   6840
      X2              =   6840
      Y1              =   3240
      Y2              =   3840
   End
   Begin VB.Line Line11 
      X1              =   5040
      X2              =   6840
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line10 
      X1              =   5040
      X2              =   5040
      Y1              =   3240
      Y2              =   3840
   End
   Begin VB.Line Line9 
      X1              =   120
      X2              =   4080
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line8 
      X1              =   120
      X2              =   7560
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   7560
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   7560
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line5 
      X1              =   7560
      X2              =   7560
      Y1              =   840
      Y2              =   4080
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   7560
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   120
      Y1              =   840
      Y2              =   4080
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   7560
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label18 
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
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   5760
      TabIndex        =   19
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000005&
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   18
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label16 
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
      Left            =   6360
      TabIndex        =   17
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000005&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   16
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Label14 
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
      Left            =   5880
      TabIndex        =   15
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000005&
      Caption         =   "Depart Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label12 
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
      Left            =   5880
      TabIndex        =   13
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label11 
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
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label10 
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
      Left            =   2040
      TabIndex        =   11
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000005&
      Caption         =   "Flight Comapny"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   4080
      X2              =   4080
      Y1              =   840
      Y2              =   4080
   End
   Begin VB.Label Label8 
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
      Left            =   2040
      TabIndex        =   9
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000005&
      Caption         =   "Cabin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label6 
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
      Left            =   2040
      TabIndex        =   7
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      Caption         =   "Departure Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label4 
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
      Left            =   2040
      TabIndex        =   5
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      Caption         =   "Arrival City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
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
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Departure City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Form3.Visible = False
Form8.Visible = True
End Sub

Private Sub Command2_Click()
Form2.Visible = True
Unload Me


End Sub

Private Sub Form_Load()
Command2.Picture = ImageList.ListImages(1).Picture
Command1.Picture = ImageList.ListImages(2).Picture
Label2.Caption = going_from
Label4.Caption = going_to
Label6.Caption = depart_date
Label8.Caption = cabin
Label10.Caption = flight_company
Label12.Caption = flight_no
Label14.Caption = depart_hour
Label16.Caption = depart_minute
Label18.Caption = price

End Sub

