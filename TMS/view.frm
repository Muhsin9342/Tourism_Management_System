VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form5 
   Caption         =   "VIEW"
   ClientHeight    =   5490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8745
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "view.frx":0000
   ScaleHeight     =   5490
   ScaleWidth      =   8745
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "DELETE USER"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   3
      Top             =   7680
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BACK"
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
      Left            =   6360
      TabIndex        =   2
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "VIEW"
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
      Left            =   4680
      TabIndex        =   1
      Top             =   6960
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4815
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   8493
      _Version        =   393216
      Rows            =   10
      Cols            =   7
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
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "view.frx":11D8B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1770
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer
i = 1
MSFlexGrid1.Clear

MSFlexGrid1.Row = 0
MSFlexGrid1.Col = 1
MSFlexGrid1.text = "  user_id"

MSFlexGrid1.Row = 0
MSFlexGrid1.Col = 2
MSFlexGrid1.text = " password"

MSFlexGrid1.Row = 0
MSFlexGrid1.Col = 3
MSFlexGrid1.text = "  name"

MSFlexGrid1.Row = 0
MSFlexGrid1.Col = 4
MSFlexGrid1.text = " gender"

MSFlexGrid1.Row = 0
MSFlexGrid1.Col = 5
MSFlexGrid1.text = " mobile_no"




If rst.State = 1 Then
  rst.Close
  End If
  
 rst.Open "select * from login ", con
 
  
  Do While Not rst.EOF
  MSFlexGrid1.Row = i
  MSFlexGrid1.Col = 1
  MSFlexGrid1.text = rst(0)
  MSFlexGrid1.Col = 2
  MSFlexGrid1.text = rst(1)
  MSFlexGrid1.Col = 3
  MSFlexGrid1.text = rst(2)
  MSFlexGrid1.Col = 4
  MSFlexGrid1.text = rst(3)
  MSFlexGrid1.Col = 5
  MSFlexGrid1.text = rst(4)
 
  
 
 
  i = i + 1
  rst.MoveNext
  MSFlexGrid1.Rows = i + 1
  If rst.EOF = True Then
   rst.Close
   Exit Do
  End If
  Loop
End Sub

Private Sub Command2_Click()
Unload Me
MDIForm1.Show
End Sub


Private Sub Command3_Click()
Dim i As Integer
i = MSFlexGrid1.RowSel
Form11.Text1.text = MSFlexGrid1.TextMatrix(i, 1)
Form11.Text2.text = MSFlexGrid1.TextMatrix(i, 2)
Form11.Text3.text = MSFlexGrid1.TextMatrix(i, 3)
Form11.Text4.text = MSFlexGrid1.TextMatrix(i, 4)
Form11.Text5.text = MSFlexGrid1.TextMatrix(i, 5)

End Sub
