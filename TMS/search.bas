Attribute VB_Name = "Module2"
Dim text As String
Dim counter As Integer
Dim rec As Integer
Dim i As Integer
Dim r As Integer
Public Sub serial_no()
If rst.State = 1 Then rst.Close

Form13.MSFlexGrid1.Height = 1200
Form13.MSFlexGrid1.Row = 0
Form13.MSFlexGrid1.Col = 1
Form13.MSFlexGrid1.text = "serial_no"
Form13.MSFlexGrid1.Col = 2
Form13.MSFlexGrid1.text = "name"
Form13.MSFlexGrid1.Col = 3
Form13.MSFlexGrid1.text = "Email id"
Form13.MSFlexGrid1.Col = 4
Form13.MSFlexGrid1.text = "mobile_no"
Form13.MSFlexGrid1.Col = 5
Form13.MSFlexGrid1.text = "city"
Form13.MSFlexGrid1.Col = 6
Form13.MSFlexGrid1.text = "travel_date"
Form13.MSFlexGrid1.Col = 7
Form13.MSFlexGrid1.text = "address"
counter = 0
If rst.State = 1 Then rst.Close

sql = "select serial_no from registration"
rst.Open sql, con, adOpenDynamic
rst.MoveFirst
While Not rst.EOF
If Form13.Text1.text = rst.Fields(0) Then
rst.Close

sql = "select * from registration where serial_no='" & Form13.Text1.text & "'"
rst.Open sql, con, adOpenDynamic
While Not rst.EOF 'to count the number of records in table
counter = counter + 1
rst.MoveNext
Wend
rst.MoveFirst
i = 1
rw = 2
Form13.MSFlexGrid1.Row = i
For rec = 1 To counter
Form13.MSFlexGrid1.Col = 0
Form13.MSFlexGrid1.text = i
Form13.MSFlexGrid1.Col = 1
Form13.MSFlexGrid1.text = rst(0)

text = Form13.MSFlexGrid1.TextMatrix(1, 1)
Form13.MSFlexGrid1.ColWidth(1) = Form13.TextWidth(text) + 1000

Form13.MSFlexGrid1.Col = 2
Form13.MSFlexGrid1.text = rst(1)

text = Form13.MSFlexGrid1.TextMatrix(1, 2)
Form13.MSFlexGrid1.ColWidth(2) = Form13.TextWidth(text) + 1000


Form13.MSFlexGrid1.Col = 3
Form13.MSFlexGrid1.text = rst(2)

'text = srchemp.MSFlexGrid1.TextMatrix(1, 3)
'srchemp.MSFlexGrid1.ColWidth(3) = srchemp.TextWidth(text) + 100


Form13.MSFlexGrid1.Col = 4
Form13.MSFlexGrid1.text = rst(3)

text = Form13.MSFlexGrid1.TextMatrix(1, 4)
Form13.MSFlexGrid1.ColWidth(4) = Form13.TextWidth(text) + 1000

Form13.MSFlexGrid1.Col = 5
Form13.MSFlexGrid1.text = rst(4)

text = Form13.MSFlexGrid1.TextMatrix(1, 5)
Form13.MSFlexGrid1.ColWidth(5) = Form13.TextWidth(text) + 1000

Form13.MSFlexGrid1.Col = 6
Form13.MSFlexGrid1.text = rst(5)

text = Form13.MSFlexGrid1.TextMatrix(1, 6)
Form13.MSFlexGrid1.ColWidth(6) = Form13.TextWidth(text) + 1000

Form13.MSFlexGrid1.Col = 7
Form13.MSFlexGrid1.text = rst(6)
rw = rw + 1
Form13.MSFlexGrid1.Rows = rw
i = i + 1
rst.MoveNext
Next rec
rst.Close
Form13.MSFlexGrid1.Height = i * 3255
Form13.MSFlexGrid1.Visible = True
Exit Sub
Else
rst.MoveNext
If rst.EOF Then
MsgBox "serial ID " & Form13.Text1.text & " Not found", vbCritical, "tourism management"
Form13.MSFlexGrid1.Visible = False
Exit Sub
End If

End If
Wend
End Sub

Public Sub name()
If rst.State = 1 Then rst.Close

Form13.MSFlexGrid1.Height = 1200
Form13.MSFlexGrid1.Row = 0
Form13.MSFlexGrid1.Col = 1
Form13.MSFlexGrid1.text = "serial_no"
Form13.MSFlexGrid1.Col = 2
Form13.MSFlexGrid1.text = "name"
Form13.MSFlexGrid1.Col = 3
Form13.MSFlexGrid1.text = "Email id"
Form13.MSFlexGrid1.Col = 4
Form13.MSFlexGrid1.text = "mobile_no"
Form13.MSFlexGrid1.Col = 5
Form13.MSFlexGrid1.text = "city"
Form13.MSFlexGrid1.Col = 6
Form13.MSFlexGrid1.text = "travel_date"
Form13.MSFlexGrid1.Col = 7
Form13.MSFlexGrid1.text = "address"
counter = 0
If rst.State = 1 Then rst.Close

sql = "select name from registration"
rst.Open sql, con, adOpenDynamic
rst.MoveFirst
While Not rst.EOF
If Form13.Text2.text = rst.Fields(0) Then
rst.Close

sql = "select * from registration where name='" & Form13.Text2.text & "'"
rst.Open sql, con, adOpenDynamic
While Not rst.EOF 'to count the number of records in table
counter = counter + 1
rst.MoveNext
Wend
rst.MoveFirst
i = 1
rw = 2
Form13.MSFlexGrid1.Row = i
For rec = 1 To counter
Form13.MSFlexGrid1.Col = 0
Form13.MSFlexGrid1.text = i
Form13.MSFlexGrid1.Col = 1
Form13.MSFlexGrid1.text = rst(0)

text = Form13.MSFlexGrid1.TextMatrix(1, 1)
Form13.MSFlexGrid1.ColWidth(1) = Form13.TextWidth(text) + 1000

Form13.MSFlexGrid1.Col = 2
Form13.MSFlexGrid1.text = rst(1)

text = Form13.MSFlexGrid1.TextMatrix(1, 2)
Form13.MSFlexGrid1.ColWidth(2) = Form13.TextWidth(text) + 1000


Form13.MSFlexGrid1.Col = 3
Form13.MSFlexGrid1.text = rst(2)

'text = srchemp.MSFlexGrid1.TextMatrix(1, 3)
'srchemp.MSFlexGrid1.ColWidth(3) = srchemp.TextWidth(text) + 100


Form13.MSFlexGrid1.Col = 4
Form13.MSFlexGrid1.text = rst(3)

text = Form13.MSFlexGrid1.TextMatrix(1, 4)
Form13.MSFlexGrid1.ColWidth(4) = Form13.TextWidth(text) + 1000

Form13.MSFlexGrid1.Col = 5
Form13.MSFlexGrid1.text = rst(4)

text = Form13.MSFlexGrid1.TextMatrix(1, 5)
Form13.MSFlexGrid1.ColWidth(5) = Form13.TextWidth(text) + 1000

Form13.MSFlexGrid1.Col = 6
Form13.MSFlexGrid1.text = rst(5)

text = Form13.MSFlexGrid1.TextMatrix(1, 6)
Form13.MSFlexGrid1.ColWidth(6) = Form13.TextWidth(text) + 1000

Form13.MSFlexGrid1.Col = 7
Form13.MSFlexGrid1.text = rst(6)
rw = rw + 1
Form13.MSFlexGrid1.Rows = rw
i = i + 1
rst.MoveNext
Next rec
rst.Close
Form13.MSFlexGrid1.Height = i * 3255
Form13.MSFlexGrid1.Visible = True
Exit Sub
Else
rst.MoveNext
If rst.EOF Then
MsgBox "NAME " & Form13.Text2.text & " Not found", vbCritical, "tourism management"
Form13.MSFlexGrid1.Visible = False
Exit Sub
End If

End If
Wend
End Sub
