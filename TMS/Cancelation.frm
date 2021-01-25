VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form10"
   MDIChild        =   -1  'True
   Picture         =   "Cancelation.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "PROCEED"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      TabIndex        =   18
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "CONFIRM CANCELATION"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   15720
      MaskColor       =   &H00FF0000&
      TabIndex        =   17
      Top             =   6840
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "reason"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   13680
      TabIndex        =   16
      Top             =   720
      Width           =   6375
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Some other personal reasons"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   23
         Top             =   4560
         Width           =   4215
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Change in wheather calamities or natural disaster in the destination place "
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   22
         Top             =   3360
         Width           =   5175
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Injury or illness to a family member or due to any other emergency"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   21
         Top             =   2400
         Width           =   4095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Work related obligation"
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
         Left            =   480
         TabIndex        =   20
         Top             =   1560
         Width           =   4395
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Select any one or more  reason for the cancelation of booking"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         TabIndex        =   19
         Top             =   480
         Width           =   5535
      End
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
      Left            =   9120
      TabIndex        =   7
      Top             =   1200
      Width           =   3015
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
      Left            =   9120
      TabIndex        =   6
      Top             =   2160
      Width           =   3015
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
      Left            =   9120
      TabIndex        =   5
      Top             =   3120
      Width           =   3015
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
      Left            =   9120
      TabIndex        =   4
      Top             =   4200
      Width           =   3015
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
      Left            =   9120
      TabIndex        =   3
      Top             =   5160
      Width           =   3015
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
      Left            =   9120
      TabIndex        =   2
      Top             =   6840
      Width           =   3495
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   9120
      TabIndex        =   1
      Top             =   6120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/mm/yy"
      Format          =   83623937
      CurrentDate     =   41897
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "Cancelation.frx":11D8B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1770
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
      Left            =   6000
      TabIndex        =   15
      Top             =   1200
      Width           =   2535
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
      Left            =   6000
      TabIndex        =   14
      Top             =   2160
      Width           =   2535
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
      Left            =   6000
      TabIndex        =   13
      Top             =   3240
      Width           =   2655
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
      Left            =   6000
      TabIndex        =   12
      Top             =   4200
      Width           =   2655
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
      Left            =   6000
      TabIndex        =   11
      Top             =   5160
      Width           =   2175
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
      Left            =   6000
      TabIndex        =   10
      Top             =   6000
      Width           =   2775
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
      Left            =   6000
      TabIndex        =   9
      Top             =   6960
      Width           =   2415
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
      Left            =   8400
      TabIndex        =   8
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BOOKING CANCELATION"
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
      Left            =   7080
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
com.CommandText = "delete from billing where serial_no='" & Form10.Text1.text & "'"
com.ActiveConnection = con
com.Execute
com1.CommandText = "delete from booking where serial_no='" & Form10.Text1.text & "' "
com1.ActiveConnection = con
com1.Execute
com2.CommandText = "delete from registration where serial_no='" & Form10.Text1.text & "' "
com2.ActiveConnection = con
com2.Execute
MsgBox "your booking have been cancelled 5% will be deducted from the total amount", vbInformation, "tourism management system"
Unload Me
MDIForm1.Show

End Sub

Private Sub Command2_Click()
Frame1.Visible = True
Command1.Enabled = True
End Sub

Private Sub Form_Load()
Frame1.Visible = False
Command1.Enabled = False
End Sub
