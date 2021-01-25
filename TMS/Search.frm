VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form13 
   Caption         =   "Search"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5760
   LinkTopic       =   "Form13"
   MDIChild        =   -1  'True
   Picture         =   "Search.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "cancel"
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
      Left            =   17160
      TabIndex        =   11
      Top             =   5400
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3255
      Left            =   2280
      TabIndex        =   10
      Top             =   6720
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   10
      Cols            =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text2 
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
      Left            =   9840
      TabIndex        =   9
      Top             =   5400
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SEARCH"
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
      Left            =   14760
      TabIndex        =   7
      Top             =   5400
      Width           =   2175
   End
   Begin VB.TextBox Text1 
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
      Left            =   9840
      TabIndex        =   6
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "select process"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11160
      TabIndex        =   2
      Top             =   1920
      Width           =   6495
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFF80&
         Caption         =   "NAME"
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
         Left            =   3480
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Sl-NO"
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
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the Name to be searched"
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
      Left            =   5040
      TabIndex        =   8
      Top             =   5400
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the Sl_no to be searched"
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
      Left            =   5040
      TabIndex        =   5
      Top             =   4080
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select anyone option to search"
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
      Left            =   5760
      TabIndex        =   1
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH"
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
      Left            =   8280
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
connect
If Option1.Value = True Then
Module2.serial_no
ElseIf Option2.Value = True Then
Module2.name
End If

End Sub


Private Sub Command2_Click()
Unload Me
MDIForm1.Show
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Text2.Visible = False
Text1.Visible = True
Label3.Visible = True
Label4.Visible = False
Text1.text = ""
Text2.text = ""
Text1.SetFocus
 MSFlexGrid1.Visible = False
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Text2.Visible = True
Text1.Visible = False
Label3.Visible = False
Label4.Visible = True
Text1.text = ""
Text2.text = ""
Text2.SetFocus
 MSFlexGrid1.Visible = False
End If
End Sub
Private Sub text2_keypress(KeyAscii As Integer)

If (KeyAscii >= 65) And (KeyAscii <= 90) Or (KeyAscii >= 97) And (KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Then
Else
KeyAscii = 0
MsgBox "No special characters or numbers allowed in name", vbCritical, "Error"
End If

End Sub

