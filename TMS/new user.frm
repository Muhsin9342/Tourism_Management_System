VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H8000000B&
   Caption         =   "NEW USER"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "new user.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "SUBMIT"
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
      Left            =   15960
      TabIndex        =   17
      Top             =   7560
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFF80&
      Caption         =   "FEMALE"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   12000
      TabIndex        =   16
      Top             =   3840
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFF80&
      Caption         =   "MALE"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   10080
      MaskColor       =   &H00FFFF80&
      TabIndex        =   15
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "BACK"
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
      Left            =   1440
      TabIndex        =   13
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLEAR"
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
      Left            =   14280
      TabIndex        =   12
      Top             =   7560
      Width           =   1455
   End
   Begin VB.TextBox Text5 
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
      Left            =   10200
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   7560
      Width           =   2535
   End
   Begin VB.TextBox Text4 
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
      Left            =   10200
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   6600
      Width           =   2535
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
      Left            =   10200
      TabIndex        =   9
      Top             =   5640
      Width           =   2535
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
      Left            =   10200
      TabIndex        =   7
      Top             =   4800
      Width           =   2535
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
      Left            =   10200
      TabIndex        =   6
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "GENDER"
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
      Left            =   6240
      TabIndex        =   14
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   240
      Picture         =   "new user.frx":11D8B
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1770
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "+91"
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
      Left            =   9600
      TabIndex        =   8
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "CONFIRM PASSWORD"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5880
      TabIndex        =   5
      Top             =   7560
      Width           =   3615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
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
      Left            =   6120
      TabIndex        =   4
      Top             =   6720
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "USER ID"
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
      Left            =   6240
      TabIndex        =   3
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "MOBILE NO"
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
      Left            =   6240
      TabIndex        =   2
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
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
      Left            =   6360
      TabIndex        =   1
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NEW USER"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      TabIndex        =   0
      Top             =   1200
      Width           =   4335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.text = ""
Option1 = False
Option2 = False
Text2.text = ""
Text3.text = ""
Text4.text = ""
Text5.text = ""
End Sub

Private Sub Command2_Click()
If Text4.text = Text5.text Then
 rst.Open "select * from login", con
   While rst.EOF <> True
     If rst(0) = Text3.text Then
        MsgBox ("User ID already exists, please try form another User ID")
        rst.Close
        Text3.text = ""
        Text4.text = ""
        Text5.text = ""
        Exit Sub
     Else
        rst.MoveNext
     End If
   Wend
   rst.Close
    Dim i As Integer
   i = 0
   If Text1.text = "" Or Text2.text = "" Or Text3.text = "" Or Text4.text = "" Or Text5.text = "" Then
       MsgBox ("Please enter all the details, text box should not be left blank"), vbInformation
     Else
     If Option1 = True Then
        com.CommandText = "insert into login values('" & Text3.text & "','" & Text5.text & "','" & Text1.text & "','" & Option1.Caption & "','" & Text2.text & "')"
     Else
        com.CommandText = "insert into login values('" & Text3.text & "','" & Text5.text & "','" & Text1.text & "','" & Option2.Caption & "','" & Text2.text & "')"
     End If
    com.ActiveConnection = con
    com.Execute
    MsgBox ("Record inserted succcessfull"), vbInformation
    Text1.text = ""
    Text2.text = ""
    Option1 = False
    Option2 = False
    Text3.text = ""
    Text4.text = ""
    Text5.text = ""
    End If
Else
  Label13.Caption = "Confirm the password correctly"
  Text4.text = ""
  Text5.text = ""
    
    
   End If
End Sub

Private Sub Command3_Click()
Unload Me
MDIForm1.Show

End Sub

Private Sub Command4_Click()
If Text4.text = Text5.text Then
 rst.Open "select * from login", con
   While rst.EOF <> True
     If rst(0) = Text3.text Then
        MsgBox ("User ID already exists, please try form another User ID")
        rst.Close
        Text3.text = ""
        Text4.text = ""
        Text5.text = ""
        Exit Sub
     Else
        rst.MoveNext
     End If
   Wend
   rst.Close
    Dim i As Integer
   i = 0
   If Text1.text = "" Or Text2.text = "" Or Text3.text = "" Or Text4.text = "" Or Text5.text = "" Then
       MsgBox ("Please enter all the details, text box should not be left blank"), vbInformation
     Else
     If Option1 = True Then
        com.CommandText = "insert into login values('" & Text3.text & "','" & Text5.text & "','" & Text1.text & "','" & Option1.Caption & "','" & Text2.text & "')"
     Else
        com.CommandText = "insert into login values('" & Text3.text & "','" & Text5.text & "','" & Text1.text & "','" & Option2.Caption & "','" & Text2.text & "')"
     End If
    com.ActiveConnection = con
    com.Execute
    MsgBox ("Record inserted succcessfull"), vbInformation
    Text1.text = ""
    Text2.text = ""
    Option1 = False
    Option2 = False
    Text3.text = ""
    Text4.text = ""
    Text5.text = ""
    End If
Else
  Label13.Caption = "Confirm the password correctly"
  Text4.text = ""
  Text5.text = ""
    End If
    
End Sub
