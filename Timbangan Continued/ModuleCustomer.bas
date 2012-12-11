Attribute VB_Name = "ModuleCustomer"
Public con As New ADODB.Connection
Public custom As New ADODB.Recordset

Public Sub konektabelcustomer()

    Set con = New ADODB.Connection
    con.Open "Provider=MSDASQL.1;Persist Security Info=False;User ID=toyib;Data Source=csc"
    con.CursorLocation = adUseClient
    
End Sub

Public Sub selekcstm()

Set custom = New ADODB.Recordset
custom.Open "SELECT * FROM customer", con, adOpenDynamic, adLockOptimistic

End Sub
