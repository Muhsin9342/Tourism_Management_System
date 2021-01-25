VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "WELCOME"
   ClientHeight    =   10380
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18330
   FillColor       =   &H00C000C0&
   BeginProperty Font 
      Name            =   "Segoe Script"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Picture         =   "welcm.frx":0000
   ScaleHeight     =   10380
   ScaleWidth      =   18330
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   840
      Top             =   1320
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ENTER"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16680
      MaskColor       =   &H00FFFFFF&
      Picture         =   "welcm.frx":BB8042
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9360
      Width           =   1455
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome To Tourism Management system"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   14295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim strspeech As String
strspeech = lblWelcome.Caption
Set objspeech = CreateObject("SAPI.spVoice")
objspeech.speak strspeech
Unload Me
FORM1.Show
End Sub
 
Private Sub Timer1_Timer()
lblWelcome.Left = lblWelcome.Left + 30
End Sub
