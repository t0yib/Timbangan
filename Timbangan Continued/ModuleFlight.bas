Attribute VB_Name = "ModuleFlight"
Public flg As New ADODB.Connection
Public flt As New ADODB.Recordset

Public Sub konekflight()

    Set flg = New ADODB.Connection
    flg.Open "Provider=MSDASQL.1;Persist Security Info=False;User ID=toyib;Data Source=csc"
    flg.CursorLocation = adUseClient
    
End Sub

Public Sub fligt()

    Set flt = New ADODB.Recordset
    flt.Open "SELECT flightcompany FROM flight order by flightcompany", flg

End Sub
