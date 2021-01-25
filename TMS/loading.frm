VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5310
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "loading.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "loading.frx":000C
   ScaleHeight     =   5310
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   480
      Top             =   2160
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Left            =   480
      Picture         =   "loading.frx":B20D
      ScaleHeight     =   315
      ScaleWidth      =   7935
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   8000
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   480
      Picture         =   "loading.frx":4C961
      ScaleHeight     =   315
      ScaleWidth      =   735
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   6480
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   480
      Top             =   1560
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Warning : Illegal use of this Software may destroy all your data. Don't Take it Easily."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   480
      TabIndex        =   5
      Top             =   4920
      Width           =   7695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading.............."
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   3360
      Width           =   2535
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 Dim i As Integer
 Dim X As Integer

Private Sub Form_Load()
File1.FileName = App.Path
X = File1.ListCount
End Sub

Private Sub Timer1_Timer()
Picture1.Visible = True
Picture1.Width = Picture1.Width + 100
If Picture1.Width = Picture2.Width Then
Timer1.Enabled = False
Unload Me
Form3.Show
End If

End Sub


Private Sub Timer2_Timer()
If (i <= X) Then
    Label1.Caption = File1.List(i)
    i = i + 1

End If
End Sub
