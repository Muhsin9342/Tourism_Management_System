VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   Caption         =   "REGISTERATION"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10665
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Registeration.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   615
      Left            =   9720
      TabIndex        =   18
      Top             =   7080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   73662467
      CurrentDate     =   41906
   End
   Begin VB.TextBox txtDTPicker 
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      TabIndex        =   17
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NEXT"
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
      Left            =   18240
      TabIndex        =   16
      Top             =   9240
      Width           =   1575
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
      Height          =   495
      Left            =   960
      TabIndex        =   15
      Top             =   9240
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   9720
      TabIndex        =   13
      Top             =   7920
      Width           =   3495
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
      Height          =   615
      Left            =   9720
      TabIndex        =   12
      Top             =   6240
      Width           =   3495
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
      Height          =   615
      Left            =   9720
      TabIndex        =   11
      Top             =   5280
      Width           =   3495
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
      Height          =   615
      Left            =   9720
      TabIndex        =   10
      Top             =   4200
      Width           =   3495
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
      Height          =   615
      Left            =   9720
      TabIndex        =   9
      Top             =   3240
      Width           =   3495
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
      Height          =   615
      Left            =   9720
      TabIndex        =   8
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "get ur serial no "
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   615
      Left            =   13440
      TabIndex        =   22
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
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
      Left            =   14160
      TabIndex        =   21
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lbl5 
      BackStyle       =   0  'Transparent
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
      Left            =   14160
      TabIndex        =   20
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lbl4 
      BackStyle       =   0  'Transparent
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
      Left            =   14160
      TabIndex        =   19
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "Registeration.frx":11D8B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1770
   End
   Begin VB.Label lblIndiacode 
      BackStyle       =   0  'Transparent
      Caption         =   "+91"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   14
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
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
      Left            =   6600
      TabIndex        =   7
      Top             =   8040
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "travel date :"
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
      Left            =   6600
      TabIndex        =   6
      Top             =   7080
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Your City:"
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
      Left            =   6600
      TabIndex        =   5
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No :"
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
      Left            =   6600
      TabIndex        =   4
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID :"
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
      Left            =   6600
      TabIndex        =   3
      Top             =   4320
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
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
      Left            =   6600
      TabIndex        =   2
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No :"
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
      Left            =   6600
      TabIndex        =   1
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTRATION"
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
      Left            =   7920
      TabIndex        =   0
      Top             =   960
      Width           =   4095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num As Integer
Private Sub Command2_Click()
connect
Dim i As String
'If txtSlno.Text = "" Or txtName.Text = "" Or txtEmailid.Text = "" Or txtmobileno.Text = "" Or txtCity.Text = "" Or txtDTPicker.Text = "" Or txtAddress.Text = "" Then
'MsgBox ("All Fields Are Manditory! Please Re-enter"), vbCritical, "Error"
If Text2.text = "" Then
MsgBox ("All Fields Are Manditory! Please enter the Name"), vbCritical, "Error"
Text2.SetFocus
ElseIf Text3.text = "" Then
MsgBox ("All Fields Are Manditory! Please enter the Email_id"), vbCritical, "Error"
Text3.SetFocus
ElseIf Text4.text = "" Then
MsgBox ("All Fields Are Manditory! Please enter the Mobile_no"), vbCritical, "Error"
Text4.SetFocus
ElseIf Text5.text = "" Then
MsgBox ("All Fields Are Manditory! Please enter the City"), vbCritical, "Error"
Text5.SetFocus
ElseIf txtDTPicker.text = "" Then
MsgBox ("All Fields Are Manditory! Please select the Travel_date"), vbCritical, "Error"
txtDTPicker.SetFocus
ElseIf Text7.text = "" Then
MsgBox ("All Fields Are Manditory! Please enter the Address"), vbCritical, "Error"
Text7.SetFocus
Else
form6.Show
form6.Text1.text = Text1.text
Me.Hide
End If
End Sub

Private Sub DTPicker1_Change()
txtDTPicker.text = Format(DTPicker1.Value, "dd-mmm-yyyy")
DTPicker1.MinDate = Date
End Sub

Private Sub Form_Load()
connect
DTPicker1.Value = Date
DTPicker1.MinDate = Date
End Sub

Private Sub Label9_Click()
Call auto
Text1.text = Format(num, "T-00")
Text2.SetFocus
End Sub

Private Sub txtDTPicker_KeyPress(KeyAscii As Integer)
KeyAscii = 0
MsgBox ("you are not allowed to enter something here"), vbCritical, "Error"
Text1.SetFocus
End Sub
Private Sub text3_KeyPress(KeyAscii As Integer)
If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 46) Or (KeyAscii = 95) Or (KeyAscii = 64) Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8)) Then
Else
KeyAscii = 0
MsgBox ("Enter the valid Email_id an Email_id can contain only Alphabets,numbers,periods(.) and special symbopls such as @ and _ "), vbCritical, "Error"
End If
End Sub
Private Sub text4_KeyPress(KeyAscii As Integer)
If ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8)) Then
Else
KeyAscii = 0
MsgBox ("Enter only Digits"), vbCritical, "Error"
End If
Text4.MaxLength = 10
End Sub
Private Sub text2_keypress(KeyAscii As Integer)
If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 32) Or (KeyAscii = 8)) Then
Else
KeyAscii = 0
MsgBox ("Enter only Alphabets"), vbCritical, "Error"
End If
End Sub
Private Sub Text1_keypress(KeyAscii As Integer)
KeyAscii = 0
MsgBox ("you are not allowed to enter something here,plz get your serial_key"), vbCritical, "Error"
Text1.SetFocus
End Sub
Private Sub text5_KeyPress(KeyAscii As Integer)
If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 32) Or (KeyAscii = 8)) Then
Else
KeyAscii = 0
MsgBox ("Enter only Alphabets"), vbCritical, "Error"
End If
End Sub

Public Sub auto()
If rst.State = 1 Then rst.Close
rst.Open "select * from registration", con
num = 1
If rst.EOF = True Then
Text1.text = Format(num, "T-01")
Else
While rst.EOF <> True
num = num + 1
rst.MoveNext
Wend
Text1.text = Format(Text1, "T-01")
End If
End Sub
