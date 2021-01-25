VERSION 5.00
Begin VB.Form Form14 
   Caption         =   "CHANGE PASSWORD"
   ClientHeight    =   8910
   ClientLeft      =   240
   ClientTop       =   900
   ClientWidth     =   14490
   LinkTopic       =   "Form14"
   MDIChild        =   -1  'True
   Picture         =   "change password.frx":0000
   ScaleHeight     =   8910
   ScaleWidth      =   14490
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFF80&
      Caption         =   "SHOW PASSWORD"
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
      Left            =   9360
      TabIndex        =   13
      Top             =   7080
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   12
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   11
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CHANGE PASSWORD"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   10
      Top             =   7920
      Width           =   1695
   End
   Begin VB.TextBox Text3 
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
      IMEMode         =   3  'DISABLE
      Left            =   9120
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   5640
      Width           =   2655
   End
   Begin VB.TextBox Text2 
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
      IMEMode         =   3  'DISABLE
      Left            =   9120
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   4680
      Width           =   2655
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
      Left            =   9120
      TabIndex        =   7
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "change password.frx":11D8B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1770
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9000
      TabIndex        =   6
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CHANGE PASSWORD"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   735
      Left            =   7320
      TabIndex        =   5
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "USER ID"
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
      Left            =   6600
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER NEW PASSWORD"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   3
      Top             =   4680
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CONFIRM PASSWORD"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      TabIndex        =   2
      Top             =   5520
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Top             =   6000
      Width           =   3495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9240
      TabIndex        =   0
      Top             =   3960
      Width           =   2415
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1 = False Then
   Text2.PasswordChar = "*"
   Text3.PasswordChar = "*"
Else
   Text2.PasswordChar = ""
   Text3.PasswordChar = ""
End If
End Sub

Private Sub Command1_Click()
If Text2.Text = "" Then
    Label7.Caption = "Enter the new password and confirm"
Else
 If Text2.Text = Text3.Text Then
    com.CommandText = "update login set password='" & Text3.Text & "' where user_id='" & Text1.Text & "'"
    com.ActiveConnection = con
    com.Execute
    MsgBox ("Your password is successfull changed, The next time you loggin this password will be used"), vbInformation
    If Text1.Text = "muhsin" Then
      Unload Me
      MDIForm1.Show
      MDIForm1.mnucancelation.Enabled = True
      MDIForm1.mnudeleteuser.Enabled = True
      MDIForm1.mnuhome.Enabled = True
      MDIForm1.mnufile.Enabled = True
      MDIForm1.mnumanipulations = True
      MDIForm1.mnuabout.Enabled = True
      MDIForm1.mnulogin.Enabled = False
      MDIForm1.mnuuserprofile.Enabled = True
      MDIForm1.mnuname.Caption = "ADMIN"
      Else
      MDIForm1.mnuname.Caption = Form14.Text1.Text
      Unload Me
      MDIForm1.Show
      MDIForm1.mnufile.Enabled = True
      MDIForm1.mnulogin.Enabled = False
      MDIForm1.mnuabout.Enabled = True
      MDIForm1.mnuhome.Enabled = False
      MDIForm1.mnudeleteuser.Enabled = False
      MDIForm1.mnucancelation.Enabled = True
      MDIForm1.mnuuserprofile.Enabled = True
      End If
Else
   Label7.Caption = "Confirm the password correctly"
   Text2.Text = ""
   Text3.Text = ""
End If
End If
End Sub

Private Sub Command2_Click()
Text2.Text = ""
Text3.Text = ""
End Sub


Private Sub Command3_Click()
If Text1.Text = "ADMIN" Then
      Unload Me
      MDIForm1.Show
      MDIForm1.mnucancelation.Enabled = True
      MDIForm1.mnudeleteuser.Enabled = True
      MDIForm1.mnuhome.Enabled = True
      MDIForm1.mnufile.Enabled = True
      MDIForm1.mnumanipulations = True
      MDIForm1.mnuabout.Enabled = True
      MDIForm1.mnulogin.Enabled = False
      MDIForm1.mnuuserprofile.Enabled = True
      MDIForm1.mnuname.Caption = "ADMIN"
      Else
      MDIForm1.mnuname.Caption = Form14.Text1.Text
      Unload Me
      MDIForm1.Show
      MDIForm1.mnufile.Enabled = True
      MDIForm1.mnulogin.Enabled = False
      MDIForm1.mnuabout.Enabled = True
      MDIForm1.mnuhome.Enabled = False
      MDIForm1.mnudeleteuser.Enabled = False
      MDIForm1.mnucancelation.Enabled = True
      MDIForm1.mnuuserprofile.Enabled = True
      
      End If
End Sub
