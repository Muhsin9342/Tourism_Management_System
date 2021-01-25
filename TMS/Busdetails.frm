VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   Picture         =   "Busdetails.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   Begin VB.CommandButton Command1 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Timer Timer6 
      Interval        =   100
      Left            =   10200
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   9120
      Top             =   360
   End
   Begin VB.Timer Timer5 
      Interval        =   2000
      Left            =   3720
      Top             =   720
   End
   Begin VB.Timer Timer4 
      Interval        =   2000
      Left            =   3240
      Top             =   720
   End
   Begin VB.Timer Timer3 
      Interval        =   2000
      Left            =   2880
      Top             =   720
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   2400
      Top             =   720
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Image Image6 
      Height          =   2415
      Left            =   16680
      Picture         =   "Busdetails.frx":11D8B
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   3405
   End
   Begin VB.Image Image5 
      Height          =   2040
      Left            =   12240
      Picture         =   "Busdetails.frx":1396B
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3555
   End
   Begin VB.Image Image4 
      Height          =   2175
      Left            =   16800
      Picture         =   "Busdetails.frx":153E8
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3255
   End
   Begin VB.Image Image3 
      Height          =   2520
      Left            =   12480
      Picture         =   "Busdetails.frx":17EA2
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   3540
   End
   Begin VB.Image Image2 
      Height          =   8775
      Left            =   600
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   11535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "LOCATION"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "Busdetails.frx":1A2A4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1770
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
form6.Show
form6.WindowState = 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyN
Image3.Top = Image3.Top - 50
Case vbKeyA
Image3.Left = Image3.Left - 50
Case vbKeyS
Image3.Top = Image3.Top + 50
Case vbKeyD
Image3.Left = Image3.Left + 50
End Select
End Sub
Private Sub Form_Load()
Text1.text = form6.Combo2.text
End Sub


Private Sub Timer1_Timer()
Image3.Top = Image3.Top - 50
Image6.Top = Image6.Top + 50

End Sub

Private Sub Timer2_Timer()
If Timer2.Enabled = True Then
a = ("C:\Users\Muhsin\Desktop\pics\train4.jpg")
Image2.Picture = LoadPicture(a)
Timer2.Enabled = False
Timer3.Enabled = True
End If

End Sub
Private Sub Timer3_Timer()
If Timer3.Enabled = True Then
b = ("C:\Users\Muhsin\Desktop\pics\train.jpg")
Image2.Picture = LoadPicture(b)
Timer3.Enabled = False
Timer4.Enabled = True
End If
End Sub
Private Sub Timer4_Timer()
If Timer4.Enabled = True Then
c = ("C:\Users\Muhsin\Desktop\pics\train1.jpg")
Image2.Picture = LoadPicture(c)
Timer4.Enabled = False
Timer5.Enabled = True
End If
End Sub


Private Sub Timer5_Timer()
If Timer5.Enabled = True Then
d = ("C:\Users\Muhsin\Desktop\pics\bus1.jpg")
Image2.Picture = LoadPicture(d)
Timer4.Enabled = False
Timer2.Enabled = True
End If
End Sub


Private Sub Timer6_Timer()
Image4.Top = Image4.Top + 50
Image5.Top = Image5.Top - 50
End Sub
