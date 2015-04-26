VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form7 
   BackColor       =   &H80000005&
   Caption         =   "Ticket"
   ClientHeight    =   5205
   ClientLeft      =   2160
   ClientTop       =   2730
   ClientWidth     =   9720
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   9720
   Begin MSComctlLib.ImageList ImageList 
      Left            =   2400
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   100
      ImageHeight     =   264
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "printticket.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "printticket.frx":1DDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "printticket.frx":3436
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2880
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   480
      Picture         =   "printticket.frx":5408
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   7680
      Picture         =   "printticket.frx":6A52
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4560
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   3975
      Left            =   0
      ScaleHeight     =   3915
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.Label Label22 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   3840
         TabIndex        =   22
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label Label21 
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
         Left            =   3720
         TabIndex        =   21
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label20 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   3480
         TabIndex        =   20
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000005&
         Caption         =   "Depart Hour"
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
         Left            =   1560
         TabIndex        =   19
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   7440
         TabIndex        =   18
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000005&
         Caption         =   "Rs"
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
         Left            =   6240
         TabIndex        =   17
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   6840
         TabIndex        =   16
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Line Line4 
         X1              =   6000
         X2              =   9600
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line3 
         X1              =   6000
         X2              =   9600
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line2 
         X1              =   6000
         X2              =   9600
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label15 
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
         Height          =   375
         Left            =   6120
         TabIndex        =   15
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   8160
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   7080
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000005&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   6000
         X2              =   6000
         Y1              =   0
         Y2              =   3960
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   3480
         TabIndex        =   11
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label10 
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
         Height          =   495
         Left            =   1560
         TabIndex        =   10
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   3480
         TabIndex        =   9
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label8 
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
         Left            =   1560
         TabIndex        =   8
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   3480
         TabIndex        =   7
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label6 
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
         Left            =   1560
         TabIndex        =   6
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   4080
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "Travelling Ticket"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   1680
         TabIndex        =   1
         Top             =   0
         Width           =   3855
      End
      Begin VB.Image Image1 
         Height          =   3960
         Left            =   0
         Picture         =   "printticket.frx":8A14
         Top             =   0
         Width           =   1500
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As Integer
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
CommonDialog1.PrinterDefault = True
CommonDialog1.ShowPrinter
Print Picture1



End Sub

Private Sub Command2_Click()
db.Close
'Unload Form1
Unload Form2
Unload Form3
Unload Form8
Form1.Visible = True
Unload Form7

End Sub

Private Sub Form_Load()
db.ConnectionString = "dsn=airlines_data;uid=" & user_name & ";pwd=" & pass_word & " "
db.Open
Image1.Picture = ImageList.ListImages(1).Picture
Command1.Picture = ImageList.ListImages(3).Picture
Command2.Picture = ImageList.ListImages(2).Picture

Label3.Caption = going_from
Label5.Caption = going_to
Label7.Caption = depart_date
Label9.Caption = cabin
Label11.Caption = flight_no
Label13.Caption = first_name
Label14.Caption = last_name
Label18.Caption = pnr_no
Label16.Caption = price
Label20.Caption = depart_hour
Label22.Caption = depart_minute

If dup_ticket = "true" Then
dup_ticket = "false"
Exit Sub
End If



'/////////////////////////////////////////////
'***************************************
'code for decrement in database about seat





'db.ConnectionString = "dsn=airlines_data;uid=system;pwd=harish"

rs.Open "select * from flight_data", db, adOpenDynamic, adLockOptimistic, adCmdText
rs.MoveFirst



Do While (Not rs.EOF)
If 0 = StrComp(flight_company, rs("flight_company").Value) Then
 If 0 = StrComp(going_from, rs("depart_city").Value) Then
        If 0 = StrComp(going_to, rs("arrival_city").Value) Then
              If 0 = StrComp(depart_date, rs("depart_date").Value) Then
                If 0 = StrComp(flight_no, rs("flight_no").Value) Then
                    If (0 = StrComp(cabin, "economy")) And (rs("eco_no_of_seat").Value >= 1) Then
                          
                         temp = rs("eco_no_of_seat").Value
                         temp = temp - 1
                         rs.Fields("eco_no_of_seat") = temp
                                       
                         rs.Update
                          Exit Sub
                                    
                    Else
                       
                       If (0 = StrComp(cabin, "business")) And (rs("busi_no_of_seat").Value >= 1) Then
                         temp = rs("busi_no_of_seat").Value
                         temp = temp - 1
                         rs.Fields("busi_no_of_seat") = temp
                         rs.Update
                         Exit Sub
                        End If
                        
                        
                    End If
                  Else 'recently added
                  rs.MoveNext
                  End If
                    
                    
                    
                    '''
                   Else
                   rs.MoveNext
                 
                            
                   End If
        Else
             
             rs.MoveNext
                    
         End If
         
Else
 rs.MoveNext
 

End If
Else
rs.MoveNext
End If

Loop



'db.Close


End Sub

