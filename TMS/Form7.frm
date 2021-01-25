VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Final Total    :"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form7"
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Calculator"
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
      Left            =   4320
      TabIndex        =   38
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "BILL"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   20295
      Begin VB.Label Label36 
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   18360
         TabIndex        =   36
         Top             =   5640
         Width           =   1335
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
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
         Left            =   18480
         TabIndex        =   35
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   18480
         TabIndex        =   34
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Vat (5.00 %)"
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
         Left            =   16920
         TabIndex        =   33
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
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
         Left            =   18360
         TabIndex        =   32
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   17160
         TabIndex        =   31
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   17280
         TabIndex        =   30
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13440
         TabIndex        =   29
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11760
         TabIndex        =   28
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10920
         TabIndex        =   27
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9720
         TabIndex        =   26
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8520
         TabIndex        =   25
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
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
         Left            =   6480
         TabIndex        =   24
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
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
         Left            =   4800
         TabIndex        =   23
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3120
         TabIndex        =   22
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
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
         Left            =   1440
         TabIndex        =   21
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   20
         Top             =   2040
         Width           =   735
      End
      Begin VB.Line Line21 
         BorderWidth     =   2
         X1              =   18240
         X2              =   19800
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Final Total :"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   16080
         TabIndex        =   19
         Top             =   5640
         Width           =   1935
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   18360
         TabIndex        =   18
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Total Members"
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
         Left            =   17040
         TabIndex        =   17
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Senior Citizens"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   15360
         TabIndex        =   16
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Childrens"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13320
         TabIndex        =   15
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Adults"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11880
         TabIndex        =   14
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12360
         TabIndex        =   13
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Seniors"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10920
         TabIndex        =   12
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Childrens"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9720
         TabIndex        =   11
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Adults"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8520
         TabIndex        =   10
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Members"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8880
         TabIndex        =   9
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Place"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Package Type"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "No_Of_Days"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Slno"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   3
         Top             =   5760
         Width           =   2775
      End
      Begin VB.Label Label11 
         Caption         =   " Date of Travel :"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   5760
         Width           =   2895
      End
      Begin VB.Line Line20 
         BorderWidth     =   2
         X1              =   360
         X2              =   19800
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line19 
         BorderWidth     =   2
         X1              =   19800
         X2              =   19800
         Y1              =   720
         Y2              =   6240
      End
      Begin VB.Line Line18 
         BorderWidth     =   2
         X1              =   18240
         X2              =   18240
         Y1              =   720
         Y2              =   6240
      End
      Begin VB.Line Line17 
         BorderWidth     =   2
         X1              =   16920
         X2              =   16920
         Y1              =   720
         Y2              =   5400
      End
      Begin VB.Line Line16 
         BorderWidth     =   2
         X1              =   15120
         X2              =   15120
         Y1              =   1320
         Y2              =   5400
      End
      Begin VB.Line Line15 
         BorderWidth     =   2
         X1              =   13200
         X2              =   13200
         Y1              =   1320
         Y2              =   5400
      End
      Begin VB.Line Line14 
         BorderWidth     =   2
         X1              =   11640
         X2              =   16920
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line12 
         BorderWidth     =   2
         X1              =   10800
         X2              =   10800
         Y1              =   1320
         Y2              =   5400
      End
      Begin VB.Line Line11 
         BorderWidth     =   2
         X1              =   9600
         X2              =   9600
         Y1              =   1320
         Y2              =   5400
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   8400
         X2              =   11640
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line9 
         BorderWidth     =   2
         X1              =   11640
         X2              =   11640
         Y1              =   720
         Y2              =   5400
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   8400
         X2              =   8400
         Y1              =   720
         Y2              =   5400
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   6360
         X2              =   6360
         Y1              =   720
         Y2              =   5400
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   4680
         X2              =   4680
         Y1              =   720
         Y2              =   5400
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   3000
         X2              =   3000
         Y1              =   720
         Y2              =   5400
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   1320
         X2              =   1320
         Y1              =   720
         Y2              =   5400
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   360
         X2              =   19800
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   360
         X2              =   360
         Y1              =   720
         Y2              =   5400
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   360
         X2              =   19800
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Mode is only by Cash"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16320
      TabIndex        =   42
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "For Senior Citizens Discount = 10%"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16320
      TabIndex        =   41
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "For Childrens Discount = 20%"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16320
      TabIndex        =   40
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "Note :"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14640
      TabIndex        =   39
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Need A Calculator? Click here :"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   37
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   89
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   88
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   87
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   86
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   85
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   84
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   83
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   82
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   81
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   80
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   79
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   78
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   77
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   76
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   75
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   74
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   73
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   72
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   71
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   70
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   69
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   68
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   67
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   66
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   65
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   64
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   63
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   62
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   61
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   60
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   59
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   58
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   57
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   56
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   55
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   54
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   53
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   52
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   51
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   50
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   49
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   48
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   47
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   46
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   45
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   44
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   43
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   42
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   41
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   40
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   39
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   38
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   37
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   36
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   35
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   34
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   33
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   32
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   31
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   30
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   29
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   28
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   27
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   26
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   25
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   24
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   23
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   22
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   21
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   20
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   19
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   18
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   17
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   16
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   15
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   14
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   13
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   12
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   11
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   10
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   9
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   8
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   7
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   6
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   5
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   4
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   3
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   2
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   19800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BILLING"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label12_Click()

End Sub
