VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "TOURISM MANAGEMENT"
   ClientHeight    =   5820
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7545
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnufile 
      Caption         =   "&FILE"
      Begin VB.Menu newfile 
         Caption         =   "NEW         "
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuexit 
         Caption         =   "EXIT "
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnumanipulations 
      Caption         =   "&MANIPULATIONS"
      WindowList      =   -1  'True
      Begin VB.Menu mnudeletion 
         Caption         =   "DELETION"
      End
      Begin VB.Menu mnuupdate 
         Caption         =   "UPDATE"
      End
      Begin VB.Menu mnusearch 
         Caption         =   "SEARCH"
      End
   End
   Begin VB.Menu mnuhome 
      Caption         =   "&HOME"
      Begin VB.Menu mnulogin 
         Caption         =   "LOGIN"
      End
      Begin VB.Menu mnuview 
         Caption         =   "VIEW"
      End
      Begin VB.Menu mnunewuser 
         Caption         =   "CREATE NEW USER"
      End
      Begin VB.Menu mnudeleteuser 
         Caption         =   "DELETE USER"
      End
   End
   Begin VB.Menu mnuuserprofile 
      Caption         =   "&USER PROFILE"
      Begin VB.Menu mnuname 
         Caption         =   ""
      End
      Begin VB.Menu mnuchangepassword 
         Caption         =   "CHANGE PASSWORD"
      End
      Begin VB.Menu mnulogout 
         Caption         =   "LOGOUT"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&ABOUT US"
   End
   Begin VB.Menu mnucancelation 
      Caption         =   "&CANCELATION"
      Begin VB.Menu mnucancelationbooking 
         Caption         =   "BOOKING CANCELATION"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuabout_Click()
Form15.Show
End Sub




Private Sub mnucancelationbooking_Click()
connect
disp = InputBox("Enter the Serial_No for cancellation", "Searching", "Enter Serial_No", 500, 700)
rst.Open "select * from registration where SERIAL_NO='" & disp & "'", con
If rst.EOF = True And rst.BOF = True Then
    MsgBox "Record Not Found"
    
 rst.Close
   Else
    Form10.Text1.text = rst(0)
  Form10.Text2.text = rst(1)
    Form10.Text3.text = rst(2)
 Form10.Text4.text = rst(3)
   Form10.Text5.text = rst(4)
Form10.DTPicker1 = rst(5)
Form10.Text7.text = rst.Fields(6)
rst.Close
End If
End Sub

Private Sub mnuchangepassword_Click()
Form14.Show
Form14.Text1.text = MDIForm1.mnuname.Caption
End Sub

Private Sub mnudeleteuser_Click()
Form11.Show
End Sub

Private Sub mnudeletion_Click()
Dim i As Variant
Form4.Label9.Visible = False
Form7.Command4.Visible = True
Form7.Command5.Visible = False
Form7.Command7.Visible = False
Form4.Text1.Enabled = False
Form4.Text2.Enabled = False
Form4.Text3.Enabled = False
Form4.Text4.Enabled = False
Form4.DTPicker1.Enabled = False
Form4.txtDTPicker.Enabled = False
Form4.Text5.Enabled = False
Form4.Text7.Enabled = False
Form4.Command1.Enabled = False
Form4.Command2.Enabled = True
form6.Text1.Enabled = False
form6.Combo1.Enabled = False
form6.Combo2.Enabled = False
form6.Option1.Enabled = False
form6.Option2.Enabled = False
form6.Option3.Enabled = False
form6.Option4.Enabled = False
form6.Option5.Enabled = False
form6.Option6.Enabled = False
form6.Text2.Enabled = False
form6.Text3.Enabled = False
form6.Text4.Enabled = False
form6.Text5.Enabled = False
form6.Text6.Enabled = False
form6.Command2.Enabled = False
form6.Command4.Enabled = False
form6.Command5.Enabled = False
form6.Command1.Enabled = True
form6.Command3.Enabled = True
Form7.Text1.Enabled = False
Form7.Command1.Enabled = False
Form7.Command3.Enabled = False
i = InputBox("Enter the Serial_No which is to be Deleted", "Searching", "Enter Serial_No", 500, 700)
Form4.lbl5.Caption = i
On Error GoTo err
rst1.Open "select * from registration where (SERIAL_NO='" & Form4.lbl5.Caption & "')", con, 2, 3
If rst1.RecordCount <> 0 Then
Form4.WindowState = 2
Form4.Show
Form4.Text1.text = rst1.Fields(0)
Form4.Text2.text = rst1.Fields(1)
Form4.Text3.text = rst1.Fields(2)
Form4.Text4.text = rst1.Fields(3)
Form4.Text5.text = rst1.Fields(4)
Form4.txtDTPicker.text = Format(rst1.Fields(5), "dd-MMM-yyyy")
Form4.Text7.text = rst1.Fields(6)
rst1.Close
Else
MsgBox "There must be atleast one record to perform this operation! Please ADD a record and then perform further manipulations", vbCritical, "Error"
rst1.Close
End If
form6.lbl1.Caption = i
rst.Open "select * from booking where SERIAL_NO='" & form6.lbl1.Caption & "' ", con, 2, 3
If rst.RecordCount <> 0 Then
form6.Text1.text = rst.Fields(0)
form6.Combo1.text = rst.Fields(1)
form6.Combo2.text = rst.Fields(2)
Form4.lbl3.Caption = rst.Fields(3)
    If Form4.lbl3.Caption = "14" Then
        form6.Option1.Value = True
        ElseIf Form4.lbl3.Caption = "20" Then
        form6.Option2.Value = True
        ElseIf Form4.lbl3.Caption = "25" Then
        form6.Option3.Value = True
    End If
Form4.lbl4.Caption = rst.Fields(4)
    If Form4.lbl4.Caption = "Regular" Then
    form6.Option4.Value = True
    ElseIf Form4.lbl4.Caption = "Premium" Then
    form6.Option5.Value = True
    ElseIf Form4.lbl4.Caption = "Deluxe" Then
    form6.Option6.Value = True
    End If
form6.Text2.text = rst.Fields(5)
form6.Text3.text = rst.Fields(6)
form6.Text4.text = rst.Fields(7)
form6.Text5.text = rst.Fields(8)
form6.Text6.text = rst.Fields(9)
rst.Close
Else
MsgBox "There must be atleast one record to perform this operation! Please ADD a record and then perform further manipulations", vbCritical, "Error"
rst.Close
End If
MsgBox "Search successfull", vbInformation, "Success"
form6.Hide
Form7.Text1.text = Form4.Text1.text
Exit Sub
err: MsgBox ("serial no does not exists,please check the serial no and try again"), vbCritical, "Error"
Unload Me
MDIForm1.Show
If rst1.State = 1 Then
rst1.Close
End If
If rst.State = 1 Then
rst.Close
End If
End Sub


Private Sub mnuexit_Click()
End

End Sub

Private Sub mnulogin_Click()
Unload Me
FORM1.Show

End Sub

Private Sub mnulogout_Click()
FORM1.Show
End Sub

Private Sub mnunewuser_Click()
Form2.Show
End Sub

Private Sub mnuopen_Click()
Form5.Show
End Sub

Private Sub mnureport_Click()
DataReport1.Show
End Sub

Private Sub mnusearch_Click()
Form13.Show
End Sub

Private Sub mnuupdate_Click()
Form7.Command7.Visible = False
Form7.Command5.Visible = False
Form7.Command6.Visible = True
Form7.Command4.Visible = False
Form4.Text1.Enabled = True
Form4.Text2.Enabled = True
Form4.Text3.Enabled = True
Form4.Text4.Enabled = True
Form4.DTPicker1.Enabled = True
Form4.txtDTPicker.Enabled = True
Form4.Text5.Enabled = True
Form4.Text7.Enabled = True
Form4.Command1.Enabled = True
Form4.Command2.Enabled = True
form6.Text1.Enabled = True
form6.Combo1.Enabled = True
form6.Combo2.Enabled = True
form6.Option1.Enabled = True
form6.Option2.Enabled = True
form6.Option3.Enabled = True
form6.Option4.Enabled = True
form6.Option5.Enabled = True
form6.Option6.Enabled = True
form6.Text2.Enabled = True
form6.Text3.Enabled = True
form6.Text4.Enabled = True
form6.Text5.Enabled = True
form6.Text6.Enabled = True
form6.Command2.Enabled = True
form6.Command4.Enabled = True
form6.Command5.Enabled = True
form6.Command1.Enabled = True
form6.Command3.Enabled = True
Form7.Text1.Enabled = True
Form7.Command1.Enabled = True
Form7.Command3.Enabled = True
Dim i As Variant
i = InputBox("Enter the Serial_No which is to be Updated", "Searching", "Enter Serial_No", 500, 700)
Form4.lbl5.Caption = i
On Error GoTo err
rst1.Open "select * from registration where (SERIAL_NO='" & Form4.lbl5.Caption & "')", con, 2, 3
If rst1.RecordCount <> 0 Then
Form4.WindowState = 2
Form4.Show
Form4.Text1.text = rst1.Fields(0)
Form4.Text2.text = rst1.Fields(1)
Form4.Text3.text = rst1.Fields(2)
Form4.Text4.text = rst1.Fields(3)
Form4.Text5 = rst1.Fields(4)
Form4.txtDTPicker.text = Format(rst1.Fields(5), "dd-MMM-yyyy")
Form4.Text7.text = rst1.Fields(6)
rst1.Close
Else
MsgBox "There must be atleast one record to perform this operation! Please ADD a record and then perform further manipulations", vbCritical, "Error"
rst1.Close
End If
form6.lbl1.Caption = i
rst.Open "select * from booking where SERIAL_NO='" & form6.lbl1.Caption & "' ", con, 2, 3
If rst.RecordCount <> 0 Then
form6.Text1 = rst.Fields(0)
form6.Combo1.text = rst.Fields(1)
form6.Combo2.text = rst.Fields(2)
Form4.lbl3.Caption = rst.Fields(3)
    If Form4.lbl3.Caption = "14" Then
        form6.Option1.Value = True
        ElseIf Form4.lbl3.Caption = "20" Then
       form6.Option2.Value = True
        ElseIf Form4.lbl3.Caption = "25" Then
     form6.Option3.Value = True
    End If
Form4.lbl4.Caption = rst.Fields(4)
    If Form4.lbl4.Caption = "Regular" Then
    form6.Option4.Value = True
    ElseIf Form4.lbl4.Caption = "Premium" Then
  form6.Option5.Value = True
    ElseIf Form4.lbl4.Caption = "Deluxe" Then
   form6.Option6.Value = True
    End If
form6.Text2.text = rst.Fields(5)
form6.Text3.text = rst.Fields(6)
form6.Text4.text = rst.Fields(7)
form6.Text5.text = rst.Fields(8)
form6.Text6.text = rst.Fields(9)
rst.Close
Else
MsgBox "There must be atleast one record to perform this operation! Please ADD a record and then perform further manipulations", vbCritical, "Error"
rst.Close
End If
MsgBox "Search successfull", vbInformation, "Success"
form6.Combo1 = ""
form6.Combo2.text = ""
form6.Image1.Picture = LoadPicture("")
form6.Hide
Form4.Show
Exit Sub
err: MsgBox ("serial no does not exists,please check the serial no and try again"), vbCritical, "Error"
Unload Me
MDIForm1.Show
If rst1.State = 1 Then
rst1.Close
End If
If rst.State = 1 Then
rst.Close
End If
End Sub


Private Sub mnuview_Click()
Unload Me
Form5.Show

End Sub

Private Sub newfile_Click()
Form4.Show
End Sub
