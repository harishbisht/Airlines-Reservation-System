VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BackColor       =   &H80000005&
   Caption         =   "All Flight Information"
   ClientHeight    =   6765
   ClientLeft      =   2010
   ClientTop       =   1515
   ClientWidth     =   8700
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   8700
   Begin MSComctlLib.ImageList ImageList 
      Left            =   9600
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   90
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Show Flight.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Show Flight.frx":0460
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Show Flight.frx":484A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Show Flight.frx":8C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Show Flight.frx":AC06
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Show Flight.frx":C440
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Show Flight.frx":FFF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Show Flight.frx":10B35
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Show Flight.frx":110DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Show Flight.frx":11C69
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Show Flight.frx":12B33
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton next_day 
      Height          =   615
      Left            =   6120
      Picture         =   "Show Flight.frx":14357
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton previous_day 
      Height          =   615
      Left            =   3000
      Picture         =   "Show Flight.frx":18731
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   480
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Result"
      Height          =   4695
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   7935
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000009&
         DisabledPicture =   "Show Flight.frx":1CB0B
         Height          =   495
         Left            =   3720
         Picture         =   "Show Flight.frx":1EACD
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000009&
         DisabledPicture =   "Show Flight.frx":20A8F
         Height          =   495
         Left            =   6960
         Picture         =   "Show Flight.frx":222B9
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton book_now 
         BackColor       =   &H80000009&
         Height          =   495
         Left            =   5160
         Picture         =   "Show Flight.frx":23AE3
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00800000&
         BorderStyle     =   6  'Inside Solid
         X1              =   3360
         X2              =   7920
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line6 
         X1              =   3360
         X2              =   7920
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line5 
         X1              =   3360
         X2              =   7920
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line4 
         X1              =   3360
         X2              =   7920
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line3 
         X1              =   3360
         X2              =   7920
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Shape Shape2 
         Height          =   735
         Left            =   5040
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Line Line2 
         X1              =   3360
         X2              =   3360
         Y1              =   120
         Y2              =   4680
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00008000&
         X1              =   3360
         X2              =   7920
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   4560
         TabIndex        =   18
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label15 
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
         Height          =   495
         Left            =   3720
         TabIndex        =   17
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label14 
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
         Left            =   360
         TabIndex        =   16
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label13 
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
         Left            =   3720
         TabIndex        =   15
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000005&
         Caption         =   "Departure Time"
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
         Left            =   3720
         TabIndex        =   14
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   6240
         TabIndex        =   13
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label10 
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
         Left            =   6000
         TabIndex        =   12
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   6720
         TabIndex        =   11
         Top             =   4080
         Width           =   375
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000005&
         Caption         =   "Result out of"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   5280
         TabIndex        =   10
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   4920
         TabIndex        =   9
         Top             =   4080
         Width           =   375
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   5640
         TabIndex        =   8
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   5640
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   5880
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "=====>>"
         Height          =   255
         Left            =   5040
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   2415
         Left            =   120
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   480
      Picture         =   "Show Flight.frx":27685
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label19 
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
      Left            =   5640
      TabIndex        =   25
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label18 
      BackColor       =   &H80000005&
      Caption         =   "Date"
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
      Left            =   4680
      TabIndex        =   24
      Top             =   1440
      Width           =   735
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   360
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000005&
      Caption         =   "No Record Found"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3360
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim count_no As Integer
Dim val As Integer
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset


Private Sub book_now_Click()


flight_company = rs("flight_company").Value
flight_no = rs("flight_no").Value
depart_hour = rs("depart_hour").Value
depart_minute = rs("depart_minute").Value


db.Close

Form3.Visible = True
'Form2.Visible = False
Unload Me



End Sub



Private Sub Command1_Click()
rs.Close
db.Close
Form1.Visible = True
Unload Me
End Sub

Private Sub Command2_Click()
If Command3.Enabled = False Then
   Command3.Enabled = True
   End If


If val + 1 = count_no Then
 Command2.Enabled = False
 
 End If
Form2.Refresh


val = val + 1
Label7.Caption = val
rs.MoveNext
Call show_data
End Sub
Private Sub Command3_Click()
If Command2.Enabled = False Then
Command2.Enabled = True
End If



rs.MoveFirst
val = 1
Label7.Caption = val
Call show_data

Command3.Enabled = False
End Sub



Private Sub Form_Load()
Command1.Picture = ImageList.ListImages(1).Picture
previous_day.Picture = ImageList.ListImages(2).Picture
next_day.Picture = ImageList.ListImages(3).Picture
Command3.Picture = ImageList.ListImages(4).Picture
Command2.Picture = ImageList.ListImages(5).Picture
book_now.Picture = ImageList.ListImages(6).Picture




'db.ConnectionString = "dsn=airlines_data;uid=system;pwd=harish"
On Error GoTo err_h:
db.ConnectionString = "dsn=airlines_data;uid=" & user_name & ";pwd=" & pass_word & " "
db.Open
rs.Open "select * from flight_data", db, adOpenDynamic, adLockOptimistic, adCmdText
db.Properties.Refresh
rs.Properties.Refresh
 Call start_searching

Exit Sub
err_h:
 
 Unload Form2
 
End Sub
Public Function check_flight()


If 0 = StrComp(going_from, rs("depart_city").Value) Then
        If 0 = StrComp(going_to, rs("arrival_city").Value) Then
              If 0 = StrComp(depart_date, rs("depart_date").Value) Then
                  If (0 = StrComp(cabin, "economy")) And rs("eco_no_of_seat").Value >= 1 Then
                      
                       count_no = count_no + 1
                       rs.MoveNext
                       
                       Exit Function
                       
                   Else
                      
                       If (0 = StrComp(cabin, "business")) And (rs("busi_no_of_seat").Value >= 1) Then
                            count_no = count_no + 1
                        rs.MoveNext
                        Exit Function
                        End If
                        rs.MoveNext
                       Exit Function
                    End If
               Else
                   rs.MoveNext
                 
                             Exit Function
                   End If
        Else
             
             rs.MoveNext
                    Exit Function
         End If
         
Else
 rs.MoveNext
 
 Exit Function
End If
 
End Function


Public Function show_data()
Do While (Not rs.EOF)

If 0 = StrComp(going_from, rs("depart_city").Value) Then
        If 0 = StrComp(going_to, rs("arrival_city").Value) Then
              If 0 = StrComp(depart_date, rs("depart_date").Value) Then
                If (0 = StrComp(cabin, "economy")) And (rs("eco_no_of_seat").Value >= 1) Then
                          
                  Label1.Caption = rs("flight_no").Value
                  Label2.Caption = going_from
                  Label4.Caption = going_to
                  Label5.Caption = depart_date
                  Label6.Caption = rs("depart_hour").Value
                  Label11.Caption = rs("depart_minute").Value
                  Label16.Caption = cabin
                  price = rs("eco_price").Value
                  Call show_images
                  'rs.MoveNext
                  Exit Function
                    
                Else
                       
                       If (0 = StrComp(cabin, "business")) And (rs("busi_no_of_seat").Value >= 1) Then
                         Label1.Caption = rs("flight_no").Value
                            Label2.Caption = going_from
                             Label4.Caption = going_to
                            Label5.Caption = depart_date
                            Label6.Caption = rs("depart_hour").Value
                            Label11.Caption = rs("depart_minute").Value
                            Label16.Caption = cabin
                            price = rs("busi_price").Value
                            Call show_images
                            Exit Function
                        Else
                            rs.MoveNext
                         ' Exit Function
                          
                        End If
                        'rs.MoveNext
                   End If
                    
                    
                    
                    
                    '''
                   Else
                   rs.MoveNext
                 
                             'Exit Function
                   End If
        Else
             
             rs.MoveNext
                    'Exit Function
         End If
         
Else
 rs.MoveNext
 
 'Exit Function
End If



Loop
End Function



Public Function show_images()

Select Case rs("flight_company").Value

Case "Jet Airways"
Image1.Picture = ImageList.ListImages(7).Picture


Case "Indigo"

Image1.Picture = ImageList.ListImages(8).Picture

Case "KingFisher"

Image1.Picture = ImageList.ListImages(9).Picture


Case "SpiceJet"


Image1.Picture = ImageList.ListImages(10).Picture
Case "AirIndia"

Image1.Picture = ImageList.ListImages(11).Picture
End Select

End Function



Private Sub next_day_Click()
depart_date = DateAdd("d", 1, depart_date)

rs.MoveFirst




 Call start_searching
 

End Sub

Private Sub previous_day_Click()
depart_date = DateAdd("d", -1, depart_date)

rs.MoveFirst
 Call start_searching

End Sub

Public Function start_searching()
rs.MoveFirst
Command2.Enabled = True
Command3.Enabled = True
Label19.Caption = depart_date 'outside frame date

count_no = 0

 Do While (Not rs.EOF)
Call check_flight
Loop

Label7.Caption = 1
Label9.Caption = count_no

val = 1
If (count_no = 1) Then
'Command3.Visible = True
'Command2.Visible = True
Command2.Enabled = False
Command3.Enabled = False

End If

rs.MoveFirst



If count_no >= 1 Then
Frame1.Visible = True
Call show_data


Else
Frame1.Visible = False
Label17.Visible = True
End If
End Function
