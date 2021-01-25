VERSION 5.00
Begin VB.Form Form15 
   Caption         =   "Form15"
   ClientHeight    =   3030
   ClientLeft      =   195
   ClientTop       =   525
   ClientWidth     =   4560
   LinkTopic       =   "Form15"
   MDIChild        =   -1  'True
   Picture         =   "About us.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14400
      TabIndex        =   2
      Top             =   9120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   2760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "About us.frx":11D8B
      Top             =   1440
      Width           =   13215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7680
      Picture         =   "About us.frx":1257B
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABOUT US"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
MDIForm1.Show
End Sub
