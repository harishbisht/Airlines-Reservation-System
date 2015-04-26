VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   3315
   ClientLeft      =   4755
   ClientTop       =   3840
   ClientWidth     =   5640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Splash Screen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "Splash Screen.frx":000C
   ScaleHeight     =   3315
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   463
      ImageHeight     =   291
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash Screen.frx":249CDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash Screen.frx":24CC70
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash Screen.frx":24D0BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash Screen.frx":24D50C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash Screen.frx":24D95A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4920
      Top             =   2760
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1680
      Top             =   0
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   5280
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   4800
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   945
      Left            =   120
      Picture         =   "Splash Screen.frx":24DDA8
      Top             =   0
      Width           =   945
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting ..."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "              .  .  .  .  .              "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Airlines Reservation System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   4575
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Option Explicit



Private Sub Form_Initialize()
frmSplash.BackColor = RGB(206, 69, 35)
   Image1.Picture = ImageList1.ListImages(1).Picture

Image2.Picture = ImageList1.ListImages(4).Picture
Image3.Picture = ImageList1.ListImages(2).Picture
End Sub

Private Sub Form_Load()

    Timer1.Enabled = True
    Timer2.Enabled = True
    
End Sub






Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = ImageList1.ListImages(4).Picture
Image3.Picture = ImageList1.ListImages(2).Picture
End Sub

Private Sub Image2_Click()
frmSplash.WindowState = 1
Form1.WindowState = 1
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = ImageList1.ListImages(5).Picture
End Sub

Private Sub Image3_Click()
End
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = ImageList1.ListImages(3).Picture
End Sub

Private Sub Timer1_Timer()
'Form1.Visible = True
'If (Form1.Visible = True) Then
'Timer1.Enabled = False
'  If frmSplash.WindowState = 0 Then
'  Form1.WindowState = 0
'  End If
'Unload frmSplash
'Else'

'Timer1.Enabled = False
'Timer1.Enabled = True
'End If

Form1.Visible = True

'Form5.Visible = True
Form5.Show vbModal

If (Form5.Visible = True) Then
Timer1.Enabled = False
  If frmSplash.WindowState = 1 Then
  'Form5.WindowState = 0
  Form5.WindowState = 1
  Form1.WindowState = 1
  
  'Load Form1
  'Form1.WindowState = 1
  End If
Unload frmSplash
Else
Unload frmSplash

Timer1.Enabled = False
Timer1.Enabled = True
End If
Unload frmSplash
End Sub

Private Sub Timer2_Timer()
Call animation
End Sub

Public Sub animation()
'Label2.Caption = ".........................................."
'Label2.Caption = "              .  .  .  .  ."
'Label2.Caption = "."
'Label2.Caption = ".."
'Label2.Caption = "..."
'Label2.Caption = "...."
'Label2.Caption = ".........................................."
i = i + 1
 If i = 0 Then
 Label2.Caption = "  ."

 End If
 
If i = 1 Then
Label2.Caption = ".  .  ."

 End If

If i = 2 Then
Label2.Caption = ".   .   .    ."

 End If

If i = 3 Then
Label2.Caption = ".  .    .   .  ."

 End If
 If i = 4 Then
Label2.Caption = "    .  .    .  .  ."

 End If
 If i = 5 Then
Label2.Caption = ".    .   .     .   .  ."
 End If
 
 If i = 6 Then
Label2.Caption = "             .  .  .  .  ."
Timer2.Interval = 200

 End If

If i = 7 Then
Label2.Caption = "                .  .  .  .  ."

 End If
If i = 8 Then
Label2.Caption = "                  .  .  .  .  ."
Timer2.Interval = 100
 End If
 'Label2.Caption = ".........................................."
If i = 9 Then
Label2.Caption = "                   .  .  .  .       ."

 End If
If i = 10 Then
Label2.Caption = "                   .  .  .        .   ."

 End If
If i = 11 Then
Label2.Caption = "                   .  .   .     .  .    ."

 End If
If i = 12 Then
Label2.Caption = "                   .  .  .     .    .     ."

 End If
If i = 13 Then
Label2.Caption = "                       .  .       .      . ."
 End If
If i = 14 Then
Label2.Caption = "                       .      .       . .  ."

 End If
If i = 15 Then
Label2.Caption = "                       .         .   .  .  ."

 End If
If i = 16 Then
Label2.Caption = "                                  .  .  . . ."

 End If
If i = 17 Then
Label2.Caption = "                                        .....    "
i = 0
 End If



End Sub
