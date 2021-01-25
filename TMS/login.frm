VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00400040&
   Caption         =   "LOGIN"
   ClientHeight    =   10710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17280
   FillColor       =   &H00800000&
   BeginProperty Font 
      Name            =   "Segoe Script"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0FFC0&
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   Picture         =   "login.frx":0000
   ScaleHeight     =   10710
   ScaleWidth      =   17280
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1440
      Top             =   6600
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   360
      Top             =   360
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXIT"
      Height          =   615
      Left            =   11040
      TabIndex        =   7
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
      Height          =   615
      Left            =   9000
      TabIndex        =   6
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      Height          =   615
      Left            =   7080
      TabIndex        =   5
      Top             =   7800
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   9960
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   6360
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   9960
      TabIndex        =   3
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   12960
      TabIndex        =   10
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   9000
      TabIndex        =   9
      Top             =   7200
      Width           =   4215
   End
   Begin VB.Image Image7 
      Height          =   2175
      Left            =   8040
      Picture         =   "login.frx":336E2
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Image Image5 
      Height          =   2160
      Left            =   8040
      Picture         =   "login.frx":3A6E6
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   3420
   End
   Begin VB.Image Image6 
      Height          =   1350
      Left            =   1800
      Picture         =   "login.frx":3BFB7
      Top             =   0
      Width           =   2250
   End
   Begin VB.Image Image4 
      Height          =   2655
      Left            =   14760
      Picture         =   "login.frx":3D1CE
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   4215
   End
   Begin VB.Image Image3 
      Height          =   2805
      Left            =   14760
      Picture         =   "login.frx":3ECCC
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   4170
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Tourism Management system"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1095
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   18975
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   2520
      Left            =   840
      Picture         =   "login.frx":4104F
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   4005
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2640
      Left            =   840
      Picture         =   "login.frx":42B7C
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   3960
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      Height          =   495
      Left            =   7080
      TabIndex        =   2
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USER ID"
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   0
      Top             =   4320
      Width           =   2895
   End
End
Attribute VB_Name = "FORM1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Dim ikeystate As Integer
Private Sub Command1_Click()
connect
Dim i As Integer
i = 0
rst.Open "select * from login", con
While rst.EOF <> True
If rst(0) = Text1.text And rst(1) = Text2.text Then
  If Text1.text = "ADMIN" Then
     MsgBox " login successfull", vbInformation, "tourism management system"
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
      MsgBox " user login successfull", vbInformation, "tourism management system"
      MDIForm1.mnuname.Caption = FORM1.Text1.text
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
 i = 1
End If
rst.MoveNext
Wend
If i = 0 Or rst.BOF = True Then
  Beep
  Label4.Caption = "Enter the valid UserID/Password"
  Text2.text = ""
End If
rst.Close

End Sub

Private Sub Command2_Click()
Text1.text = ""
Text2.text = ""

End Sub

Private Sub Command3_Click()
End

End Sub

Private Sub Form_Load()
Label5.Visible = False
End Sub


Private Sub Text2_Change()
Call capslockon
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.ToolTipText = "Enter the password"
End Sub

Private Sub Timer1_Timer()
If Image5.Visible = True Then
Image7.Visible = True
Image5.Visible = False
ElseIf Image7.Visible = True Then
Image7.Visible = False
Image5.Visible = True
End If
End Sub

Public Function capslockon() As Boolean
Dim ikeystate As Integer
    ikeystate = GetKeyState(vbKeyCapital)
    capslockon = (ikeystate = 1 Or ikeystate = -127)
        If capslockon = True Then
Beep
Label5.Visible = True
Label5.Caption = "caps on"
ElseIf capslockon = False Then
Label5.Caption = ""

    End If
End Function
