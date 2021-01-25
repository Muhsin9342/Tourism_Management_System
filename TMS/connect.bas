Attribute VB_Name = "Module1"
Public con As ADODB.Connection
Public com As ADODB.Command
Public com1 As ADODB.Command ' this line is used because in the create user we are using two insert statements so we have created com12 object;
Public com2 As ADODB.Command
Public rst As ADODB.Recordset
Public rst1 As New ADODB.Recordset

Public Sub connect()
Set con = New ADODB.Connection
    Set com = New ADODB.Command
    Set com1 = New ADODB.Command ' this line is used because in the create user we are using two insert statements so we have created com12 object;
    Set com2 = New ADODB.Command
    Set rst = New ADODB.Recordset
   Set rst1 = New ADODB.Recordset
   ' this is used for access connection
    With con
        .ConnectionString = "User ID = system;Password = muhsin;data source = muhsin"
        
        .Provider = "MSDASQL"
        .Open
    End With
    com.ActiveConnection = con
    
End Sub




