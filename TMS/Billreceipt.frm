VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form9"
   MDIChild        =   -1  'True
   Picture         =   "Billreceipt.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
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
      Left            =   18480
      TabIndex        =   44
      Top             =   9600
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DONE"
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
      Left            =   840
      TabIndex        =   43
      Top             =   9480
      Width           =   1575
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
      Height          =   7695
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   20295
      Begin VB.Label Label47 
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
         Left            =   3600
         TabIndex        =   42
         Top             =   6960
         Width           =   2775
      End
      Begin VB.Label Label44 
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
         Left            =   3600
         TabIndex        =   41
         Top             =   6360
         Width           =   2655
      End
      Begin VB.Label Label46 
         Caption         =   "Date of arrival :"
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
         Left            =   240
         TabIndex        =   40
         Top             =   6960
         Width           =   2775
      End
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   5640
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
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
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
         Left            =   15240
         TabIndex        =   4
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         Caption         =   "Source :"
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
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   6360
         Width           =   1215
      End
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   17520
      Picture         =   "Billreceipt.frx":11D8B
      Top             =   9480
      Width           =   720
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HAVE A GREAT JOURNEY AHEAD ""TRAVEL WITH EASE"""
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
      Left            =   3480
      TabIndex        =   1
      Top             =   9600
      Width           =   10935
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "Billreceipt.frx":2015F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1770
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "BILLING RECEIPT"
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
      Left            =   8760
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form9.PrintForm
End Sub
Private Sub Command2_Click()
Unload Me
Unload Form7
MDIForm1.Show
End Sub

Private Sub Form_Load()
Label20.Caption = Form7.Label20.Caption
Label21.Caption = Form7.Label21.Caption
Label22.Caption = Form7.Label21.Caption
Label23.Caption = Form7.Label23.Caption
Label24.Caption = Form7.Label24.Caption
Label25.Caption = Form7.Label25.Caption
Label26.Caption = Form7.Label26.Caption
Label27.Caption = Form7.Label27.Caption
Label28.Caption = Form7.Label28.Caption
Label29.Caption = Form7.Label29.Caption
Label45.Caption = Form7.Label45.Caption
Label31.Caption = Form7.Label31.Caption
Label32.Caption = Form7.Label32.Caption
Label35.Caption = Form7.Label35.Caption
Label36.Caption = Form7.Label36.Caption
Label2.Caption = Form7.Label2.Caption
Label44.Caption = Form7.Label44.Caption
Label47.Caption = Form7.Label47.Caption
End Sub
