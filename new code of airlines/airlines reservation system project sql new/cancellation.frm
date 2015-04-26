VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   BackColor       =   &H80000005&
   Caption         =   "Cancellation"
   ClientHeight    =   4110
   ClientLeft      =   3555
   ClientTop       =   2670
   ClientWidth     =   7005
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7005
   Visible         =   0   'False
   Begin MSComctlLib.ImageList ImageList 
      Left            =   8160
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
            Picture         =   "cancellation.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cancellation.frx":0460
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   4920
      Picture         =   "cancellation.frx":3152
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   480
      Picture         =   "cancellation.frx":5E34
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.Line Line15 
      X1              =   4800
      X2              =   4800
      Y1              =   3120
      Y2              =   3840
   End
   Begin VB.Line Line14 
      X1              =   6600
      X2              =   6600
      Y1              =   3120
      Y2              =   3840
   End
   Begin VB.Line Line13 
      X1              =   4800
      X2              =   6600
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line12 
      X1              =   4800
      X2              =   6600
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line11 
      X1              =   360
      X2              =   1920
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line10 
      X1              =   1920
      X2              =   1920
      Y1              =   120
      Y2              =   720
   End
   Begin VB.Line Line9 
      X1              =   360
      X2              =   1920
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line8 
      X1              =   360
      X2              =   360
      Y1              =   120
      Y2              =   720
   End
   Begin VB.Line Line7 
      X1              =   360
      X2              =   3600
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line6 
      X1              =   360
      X2              =   3600
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line5 
      X1              =   360
      X2              =   3600
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line4 
      X1              =   3600
      X2              =   3600
      Y1              =   960
      Y2              =   3480
   End
   Begin VB.Line Line3 
      X1              =   360
      X2              =   3600
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   3600
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   360
      Y1              =   960
      Y2              =   3480
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000005&
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000005&
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000005&
      Caption         =   "PNR NO"
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
      Left            =   480
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      Caption         =   "Mobile No"
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
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "Last Name"
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
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "First Name"
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
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As Integer
Dim db As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim rs As New ADODB.Recordset



Private Sub Command1_Click()
db.Close

Unload Me

End Sub

Private Sub Command2_Click()
 
 cabin = rs.Fields("cabin")
 flight_no = rs.Fields("flight_no")
 Call increment_seat
   
rs.Delete

db.Close
MsgBox " booking cancelled sucessful", vbInformation
Unload Me
End Sub

Private Sub Form_Load()
Command1.Picture = ImageList.ListImages(1).Picture
Command2.Picture = ImageList.ListImages(2).Picture

On Error GoTo err_h:
'db.ConnectionString = "dsn=airlines_data;uid=system;pwd=harish"
db.ConnectionString = "dsn=airlines_data;uid=" & user_name & ";pwd=" & pass_word & " "
db.Open
rs.Open "select * from cust_data where pnr_no =  '" & pnr_cancel & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'Label1.Caption = rs("first_name").Value

'Label1.Caption = rs.Fields("first_name")

If (Not rs.EOF) Then
Form4.Visible = True
Form4.Refresh
Label5.Caption = rs.Fields("first_name")
Label6.Caption = rs.Fields("last_name")
Label7.Caption = rs.Fields("mobile_no")
Label8.Caption = rs.Fields("pnr_no")



Else
pnr_cancel = "notfound"
db.Close
End If
Exit Sub
err_h:

Unload Form4
End Sub








Public Function increment_seat()




rs1.Open "select * from flight_data where flight_no =  '" & flight_no & "'  ", db, adOpenDynamic, adLockOptimistic, adCmdText


If cabin = "economy" Then

 temp = rs1("eco_no_of_seat").Value
    
    
   temp = temp + 1
    
    rs1.Fields("eco_no_of_seat") = temp
            rs1.Update


Else
 temp = rs1("busi_no_of_seat").Value
   temp = temp + 1
    rs1.Fields("busi_no_of_seat") = temp
                                       
    rs1.Update




End If
End Function

