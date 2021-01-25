VERSION 5.00
Begin VB.Form form6 
   Caption         =   "  "
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7545
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Booking.frx":0000
   ScaleHeight     =   5715
   ScaleWidth      =   7545
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "VIEW"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      TabIndex        =   34
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "PRICE"
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
      Left            =   6960
      TabIndex        =   33
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "TOTAL MEMBER"
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
      Left            =   6960
      TabIndex        =   32
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NEXT"
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
      Left            =   18720
      TabIndex        =   31
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
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
      Left            =   480
      TabIndex        =   30
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
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
      Left            =   480
      TabIndex        =   29
      Top             =   9360
      Width           =   1215
   End
   Begin VB.TextBox Text6 
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
      Left            =   8400
      TabIndex        =   28
      Top             =   9000
      Width           =   2535
   End
   Begin VB.TextBox Text5 
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
      Left            =   8400
      TabIndex        =   27
      Top             =   8280
      Width           =   2535
   End
   Begin VB.TextBox Text4 
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
      Left            =   8400
      TabIndex        =   24
      Top             =   7560
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   8400
      TabIndex        =   23
      Top             =   6840
      Width           =   2535
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
      Height          =   495
      Left            =   8400
      TabIndex        =   22
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Package Type :"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6240
      TabIndex        =   15
      Top             =   4080
      Width           =   8295
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFF80&
         Caption         =   "Deluxe"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   18
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFF80&
         Caption         =   "Premium"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   17
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFF80&
         Caption         =   "Regular"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Booking.frx":11D8B
      Left            =   8400
      List            =   "Booking.frx":11D95
      TabIndex        =   13
      Top             =   5400
      Width           =   2475
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "No Of Days :"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6240
      TabIndex        =   8
      Top             =   2760
      Width           =   9495
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFF80&
         Caption         =   "25Days/24 Nights"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFF80&
         Caption         =   "20 Days/19 Nights"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   10
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF80&
         Caption         =   "14 Days/13 Nights"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Booking.frx":11DA5
      Left            =   13440
      List            =   "Booking.frx":11DA7
      TabIndex        =   6
      Top             =   2040
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Booking.frx":11DA9
      Left            =   5640
      List            =   "Booking.frx":11DB9
      TabIndex        =   0
      Top             =   2040
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
      Left            =   4920
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   36
      Top             =   9720
      Width           =   1215
   End
   Begin VB.Label lblprice 
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
      Left            =   8400
      TabIndex        =   35
      Top             =   9720
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   1215
      Left            =   0
      Picture         =   "Booking.frx":11DF6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1770
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "price for a single person :"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   26
      Top             =   8880
      Width           =   3375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Members  :"
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
      Left            =   2880
      TabIndex        =   25
      Top             =   8280
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "No of Senior Citizens :"
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
      Left            =   3000
      TabIndex        =   21
      Top             =   7560
      Width           =   2655
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "No of Children :"
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
      Left            =   3000
      TabIndex        =   20
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "No of Adults :"
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
      Left            =   3120
      TabIndex        =   19
      Top             =   6120
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   4695
      Left            =   12240
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   6255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Type of package :"
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
      Left            =   1560
      TabIndex        =   14
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Transport    :"
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
      Left            =   3120
      TabIndex        =   12
      Top             =   5400
      Width           =   3495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the No of days in the package :"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      TabIndex        =   7
      Top             =   3000
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Selelect the place to travel :"
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
      Left            =   9000
      TabIndex        =   5
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select The State To Travel :"
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
      Left            =   1560
      TabIndex        =   4
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sl_no :"
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
      Left            =   2040
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BOOKING DETAILS"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   1
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Date
Dim b As String
Dim c As Date
Private Sub Combo1_Click()
If Combo1.text = "DELHI" Then
Combo2.List(0) = "LAL KOT"
Combo2.List(1) = "SIRI"
Combo2.List(2) = "FIROZABAD"
ElseIf Combo1.text = "JAMMU & KASHMIR" Then
Combo2.List(0) = "GULMARG"
Combo2.List(1) = "SRINAGAR"
Combo2.List(2) = "JAMMU"
ElseIf Combo1.text = "HIMACHAL PRADESH" Then
Combo2.List(0) = "SHIMLA"
Combo2.List(1) = "MANALI"
Combo2.List(2) = "KULLU"
ElseIf Combo1.text = "UTTAR PRADESH" Then
Combo2.List(0) = "AGRA"
Combo2.List(1) = "LUCKNOW"
Combo2.List(2) = "FATHEPUR SHIKRI"
End If
End Sub
Private Sub Combo2_Click()
Dim i As String
If Combo2.text = "LAL KOT" Then
i = "C:\Users\Muhsin\Desktop\pics\lalkot.jpg"
Image1.Picture = LoadPicture(i)
lblprice.Caption = "5000"
ElseIf Combo2.text = "SIRI" Then
i = "C:\Users\Muhsin\Desktop\pics\siri.jpg"
Image1.Picture = LoadPicture(i)
lblprice.Caption = "5000"
ElseIf Combo2.text = "FIROZABAD" Then
i = "C:\Users\Muhsin\Desktop\pics\firozabad.jpg"
Image1.Picture = LoadPicture(i)
lblprice.Caption = "5000"
ElseIf Combo2.text = "GULMARG" Then
i = "C:\Users\Muhsin\Desktop\pics\gulmarg.jpg"
Image1.Picture = LoadPicture(i)
lblprice.Caption = "8000"
ElseIf Combo2.text = "SRINAGAR" Then
i = "C:\Users\Muhsin\Desktop\pics\srinagar.jpg"
Image1.Picture = LoadPicture(i)
lblprice.Caption = "8500"
ElseIf Combo2.text = "JAMMU" Then
i = "C:\Users\Muhsin\Desktop\pics\jammu.jpg"
Image1.Picture = LoadPicture(i)
lblprice.Caption = "6000"
ElseIf Combo2.text = "SHIMLA" Then
i = "C:\Users\Muhsin\Desktop\pics\shimla.jpg"
Image1.Picture = LoadPicture(i)
lblprice.Caption = "10000"
ElseIf Combo2.text = "MANALI" Then
i = "C:\Users\Muhsin\Desktop\pics\manali.jpg"
Image1.Picture = LoadPicture(i)
lblprice.Caption = "9000"
ElseIf Combo2.text = "KULLU" Then
i = "C:\Users\Muhsin\Desktop\pics\kullu.jpg"
Image1.Picture = LoadPicture(i)
lblprice.Caption = "9500"
ElseIf Combo2.text = "AGRA" Then
i = "C:\Users\Muhsin\Desktop\pics\tajmahal.jpg"
Image1.Picture = LoadPicture(i)
lblprice.Caption = "7000"
ElseIf Combo2.text = "LUCKNOW" Then
i = "C:\Users\Muhsin\Desktop\pics\fathe.jpg"
Image1.Picture = LoadPicture(i)
lblprice.Caption = "7500"
ElseIf Combo2.text = "FATHEPUR SHIKRI" Then
i = "C:\Users\Muhsin\Desktop\pics\fathepur.jpg"
Image1.Picture = LoadPicture(i)
lblprice.Caption = "8000"
    End If
End Sub
Private Sub Command1_Click()
Me.Hide
Form4.Show
End Sub
Private Sub Command2_Click()
Combo1.text = ""
Combo2.text = ""
Combo3.text = ""
Text2.text = ""
Text3.text = ""
Text4.text = ""
Text5.text = ""
Text6.text = ""
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Image1.Picture = LoadPicture("")
Combo1.SetFocus
End Sub
Private Sub Command3_Click()
If Text1.text = "" Then
MsgBox ("An unexpected error occured while the generation of Serial_no"), vbCritical, "Error"
ElseIf Combo1.text = "" Then
MsgBox ("All Fields Are Manditory! Please select a state from the dropdown list"), vbCritical, "Error"
Combo1.SetFocus
ElseIf Combo2.text = "" Then
MsgBox ("All Fields Are Manditory! Please select a place from the dropdown list"), vbCritical, "Error"
Combo2.SetFocus
ElseIf (Option1.Value = False And Option2.Value = False And Option3.Value = False) Then
MsgBox ("All Fields Are Manditory! Please Select the No_Of_Days option among the various options provided"), vbCritical, "Error"
Option1.SetFocus
ElseIf (Option4.Value = False And Option5.Value = False And Option6.Value = False) Then
MsgBox ("All Fields Are Manditory! Please select the Package_Type option among the various options provided"), vbCritical, "Error"
Option4.SetFocus
ElseIf (Text2.text = "" And Text3.text = "" And Text4.text = "") Then
MsgBox ("All Fields Are Manditory! Please enter atleast one member to continue the booking process"), vbCritical, "Error"
Text2.SetFocus

ElseIf Text5.text = "" Then
MsgBox ("All Fields Are Manditory! Please Click on Total_Members command button"), vbCritical, "Error"
Command4.SetFocus
ElseIf Text6.text = "" Then
MsgBox ("All Fields Are Manditory! Please Click on Price_Per_Person command button"), vbCritical, "Error"
Command5.SetFocus
Else
Form7.Show
Me.Hide
Form7.Text1.text = Form4.Text1.text
Form7.Label44.Caption = Form4.Text5.text
Form7.Label20.Caption = Form4.Text1.text
Form7.Label21.Caption = Form4.Text2.text
If form6.Option1.Value = True Then
Form7.Label22.Caption = "14"
ElseIf form6.Option2.Value = True Then
Form7.Label22.Caption = "20"
ElseIf form6.Option3.Value = True Then
Form7.Label22.Caption = "25"
End If
If form6.Option4.Value = True Then
Form7.Label23.Caption = "Regular"
ElseIf form6.Option5.Value = True Then
Form7.Label23.Caption = "Premium"
ElseIf form6.Option6.Value = True Then
Form7.Label23.Caption = "Deluxe"
End If
Form7.Label24.Caption = form6.Combo2.text
Form7.Label25.Caption = form6.Text2.text
Form7.Label26.Caption = form6.Text3.text
Form7.Label27.Caption = form6.Text4.text
If Val(form6.Text2.text) > 0 Then
Form7.Label28.Caption = Val(form6.Text6.text) * Val(form6.Text2.text)
Else
Form7.Label28.Caption = 0
End If
If Val(form6.Text3.text) > 0 Then
Form7.Label29.Caption = ((Val(form6.Text6.text) * Val(form6.Text3.text)) * 80 / 100)
Else
Form7.Label29.Caption = 0
End If
If Val(form6.Text4.text) > 0 Then
Form7.Label45.Caption = ((Val(form6.Text6.text) * Val(form6.Text4.text)) * 90 / 100)
Else
Form7.Label45.Caption = 0
End If
Form7.Label31.Caption = form6.Text5.text
Form7.Label32.Caption = Val(Form7.Label28.Caption) + Val(Form7.Label29.Caption) + Val(Form7.Label45.Caption)
Form7.Label35.Caption = (Val(Form7.Label32.Caption) * 5 / 100)
Form7.Label36.Caption = (Val(Form7.Label32.Caption) + Val(Form7.Label35.Caption))
Form7.Label2.Caption = Form4.txtDTPicker.text
Form7.Label44.Caption = Form4.Text5.text
a = Form4.txtDTPicker.text
b = Val(Form7.Label22.Caption)
c = a + b
Form7.Label47.Caption = Format(c, "dd-MMM-yyyy")
End If

End Sub
Private Sub Command4_Click()
If Text2.text = "" And Text3.text = "" And Text4.text = "" Then
MsgBox ("Please Re-enter the details! Atleast select one member for further booking process"), vbCritical, "Error"
Text2.SetFocus
Else
If Text2.text = "" Then
Text2.text = "0"
End If
If Text3.text = "" Then
Text3.text = "0"
End If
If Text4.text = "" Then
Text4.text = "0"
End If
Text5.text = (Val(Text2.text) + Val(Text3.text) + Val(Text4.text))
End If
End Sub
Private Sub Command5_Click()
If Combo1.text = "" Then
MsgBox "Please select the state from the dropdown list", vbCritical, "error"
Combo1.SetFocus
ElseIf Combo2.text = "" Then
MsgBox "Please select the place from the dropdown list", vbCritical, "error"
Combo2.SetFocus
ElseIf (Option1.Value = False And Option2.Value = False And Option3.Value = False) Then
MsgBox ("All Fields Are Manditory! Please Select the No_Of_Days option among the various options provided"), vbCritical, "Error"
Option1.SetFocus
ElseIf (Option4.Value = False And Option5.Value = False And Option6.Value = False) Then
MsgBox ("All Fields Are Manditory! Please select the Package_Type option among the various options provided"), vbCritical, "Error"
Option4.SetFocus
ElseIf Text2.text = "" And Text3.text = "" And Text3.text = "" Then
MsgBox "Please enter atleast one member for the further booking process", vbCritical, "error"
Text2.SetFocus
ElseIf Text5.text = "" Then
MsgBox "Please click on the Total Members command button to continue the booking process", vbCritical, "Error"
Command4.SetFocus
Else
If Option1.Value = True And Option4.Value = True Then
Text6.text = Val(lblprice.Caption)
ElseIf Option1.Value = True And Option5.Value = True Then
Text6.text = (Val(lblprice.Caption) + Val(lblprice.Caption) * 50 / 100)
ElseIf Option1.Value = True And Option6.Value = True Then
Text6.text = (Val(lblprice.Caption) + Val(lblprice.Caption))
End If
If Option2.Value = True And Option4.Value = True Then
Text6.text = (Val(lblprice.Caption) * 2)
ElseIf Option2.Value = True And Option5.Value = True Then
Text6.text = ((Val(lblprice.Caption) + Val(lblprice.Caption) * 50 / 100) * 2)
ElseIf Option2.Value = True And Option6.Value = True Then
Text6.text = ((Val(lblprice.Caption) + Val(lblprice.Caption)) * 2)
End If
If Option3.Value = True And Option4.Value = True Then
Text6.text = (Val(lblprice.Caption) * 4)
ElseIf Option3.Value = True And Option5.Value = True Then
Text6.text = ((Val(lblprice.Caption) + Val(lblprice.Caption) * 50 / 100) * 4)
ElseIf Option3.Value = True And Option6.Value = True Then
Text6.text = ((Val(lblprice.Caption) + Val(lblprice.Caption)) * 4)
End If
End If
End Sub
Private Sub Command6_Click()
Me.Hide
Form8.Show
Form8.WindowState = 2
Form8.Text1.text = form6.Combo2.text
End Sub
Private Sub Form_Load()
connect
form6.Command1.Enabled = True
form6.Command3.Enabled = True
End Sub
Private Sub combo2_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
  MsgBox ("you are not allowed to enter something here! Please select a value from the dropdown list"), vbCritical, "Error"
End Sub
Private Sub text5_KeyPress(KeyAscii As Integer)
KeyAscii = 0
MsgBox ("You are not allowed to enter something here! Please click on Total Members command button"), vbCritical, "Error"
End Sub
Private Sub text2_keypress(KeyAscii As Integer)
If ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8)) Then
Else
KeyAscii = 0
MsgBox ("Enter only Digits"), vbCritical, "Error"
End If
Text2.MaxLength = 4
End Sub
Private Sub text3_KeyPress(KeyAscii As Integer)
If ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8)) Then
Else
KeyAscii = 0
MsgBox ("Enter only Digits"), vbCritical, "Error"
End If
Text3.MaxLength = 4
End Sub
Private Sub text6_KeyPress(KeyAscii As Integer)
KeyAscii = 0
MsgBox ("You are not allowed to enter something here! Please click on Price  command button"), vbCritical, "Error"
End Sub
Private Sub text4_KeyPress(KeyAscii As Integer)
If ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8)) Then
Else
KeyAscii = 0
MsgBox ("Enter only Digits"), vbCritical, "Error"
End If
Text4.MaxLength = 4
End Sub
Private Sub Text1_keypress(KeyAscii As Integer)
KeyAscii = 0
MsgBox ("you are not allowed to enter something here"), vbCritical, "Error"
End Sub
Private Sub combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
MsgBox ("you are not allowed to enter something here! Please select a value from the dropdown list"), vbCritical, "Error"
End Sub
