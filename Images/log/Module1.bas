Attribute VB_Name = "Module1"
Public con As ADODB.Connection
Public com As ADODB.Command
Public com1 As ADODB.Command
Public rst As ADODB.Recordset
Public rst1 As ADODB.Recordset


Public Sub connect()
Set con = New ADODB.Connection
Set com = New ADODB.Command
Set com1 = New ADODB.Command
Set rst = New ADODB.Recordset
With con
'.ConnectionString = "data source=muhsin'
.ConnectionString = "user id=SYSTEM;password=muhsin;data source=muhsin"
.Provider = "MSDASQL"
.Open
End With
com.ActiveConnection = con

End Sub
