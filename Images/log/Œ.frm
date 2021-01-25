VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   7
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "newuser"
      Height          =   615
      Left            =   3000
      TabIndex        =   6
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "login"
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "login"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "password"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "username"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "enter the details"
connect
Dim i As Integer
i = o
rst.Open "select * from login", con
While rst.EOF <> True
If rst(1) = Text2.Text And rst(0) = Text1.Text Then
MsgBox "login successfull"
i = 1
Unload Me
Else
rst.MoveNext
End If
Wend
If i = 0 Or rst.BOF = True Then
MsgBox "record not found"
End If
rst.Close
End If

End Sub
