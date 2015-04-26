VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form5 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Administration Login"
   ClientHeight    =   4515
   ClientLeft      =   3300
   ClientTop       =   2880
   ClientWidth     =   6180
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList 
      Left            =   1320
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   416
      ImageHeight     =   315
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "admin_login.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "admin_login.frx":5FFF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "admin_login.frx":63764
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "admin_login.frx":64EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "admin_login.frx":66760
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      Height          =   4725
      Left            =   0
      ScaleHeight     =   8.334
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   10.927
      TabIndex        =   4
      Top             =   -30
      Width           =   6195
      Begin VB.CommandButton back 
         BackColor       =   &H80000009&
         Height          =   495
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Height          =   495
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   720
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1920
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   720
         TabIndex        =   0
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Image Image2 
         Height          =   630
         Left            =   5400
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   5400
         Top             =   1080
         Visible         =   0   'False
         Width           =   705
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub back_Click()
End
End Sub

Private Sub Command1_Click()
Dim i As Integer
i = 0
On Error GoTo err_h:
If Text1.Text = "key" And Text2.Text = "key" Then
Form1.Visible = True
Form5.Visible = False

login_uname = "key"
login_passwd = "key"

Exit Sub
End If

If rs.BOF Then
MsgBox "currently no password is added in database .. use master password", vbInformation
Exit Sub
End If

Dim user As Variant
Dim pass As Variant

'Open "adminlogin.dat" For Binary As #1
'Get #1, , user
'Get #1, , pass
'Close #1
'db.ConnectionString = "dsn=airlines_data;uid=" & user_name & ";pwd=" & pass_word & " "
'user = rs.Fields(0)
'pass = rs.Fields(1)

If Text1.Text = "" Then
Image1.Picture = ImageList.ListImages(4).Picture
Image1.Visible = True
End If

If Text2.Text = "" Then
Image2.Picture = ImageList.ListImages(4).Picture
Image2.Visible = True
End If
rs.MoveFirst

Do While Not rs.EOF
If Text1.Text = rs.Fields(0).Value Or Text1.Text = "key" Then
      Image1.Picture = ImageList.ListImages(3).Picture
      Image1.Visible = True
       i = 1
  If Text2.Text = rs.Fields(1).Value Or Text2.Text = "key" Then
       login_uname = rs.Fields(0).Value
       login_passwd = rs.Fields(1).Value
       'rs.Close
       'db.Close
      Form1.Visible = True
      
      'Form5.Visible = False
      Text1.Text = ""
      Text2.Text = ""
      Form5.Visible = False
      'Unload Form5
      Exit Sub
      'Image2.Picture = ImageList.ListImages(3).Picture
      'Image2.Visible = True
   Else
     ' MsgBox " wrong password", vbCritical
      Image2.Picture = ImageList.ListImages(4).Picture
        Image2.Visible = True
  End If
Else
    Image1.Picture = ImageList.ListImages(4).Picture
      Image1.Visible = True
 ' MsgBox "wrong user name", vbCritical
End If
rs.MoveNext
Loop
''db.Close
'db.Close


If i = 1 Then
Image1.Picture = ImageList.ListImages(3).Picture
      Image1.Visible = True
      End If
Exit Sub
err_h:
MsgBox "contect to your admin or not connected to database"
End Sub

Private Sub Form_Load()


On Error GoTo err_h:
Open "database_connectivity.dat" For Binary As #1
Get #1, , user_name
Get #1, , pass_word
Close #1

Unload frmSplash
Text1.TabIndex = 0
Picture1.Picture = ImageList.ListImages(1).Picture
Command1.Picture = ImageList.ListImages(2).Picture
'back.Picture = ImageList.ListImages(3).Picture
Image1.Picture = ImageList.ListImages(3).Picture
Image2.Picture = ImageList.ListImages(4).Picture

back.Picture = ImageList.ListImages(5).Picture

'db.ConnectionString = "dsn=airlines_data;uid=system;pwd=harish"
db.ConnectionString = "dsn=airlines_data;uid=" & user_name & ";pwd=" & pass_word & " "
db.Open
rs.Open "select * from login", db, adOpenDynamic, adLockOptimistic, adCmdText


Exit Sub
err_h:
MsgBox "currently no password is added in database .. use master password at login", vbInformation

End Sub



Private Sub Text1_GotFocus()
Image1.Visible = False
End Sub

Private Sub Text2_GotFocus()
Image2.Visible = False
End Sub
