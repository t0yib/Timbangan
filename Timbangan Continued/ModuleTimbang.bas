Attribute VB_Name = "ModuleTimbang"
Public cstm As New ADODB.Connection
Public kategori As New ADODB.Recordset

Public Sub konektimbang()

    Set cstm = New ADODB.Connection
    cstm.Open "Provider=MSDASQL.1;Persist Security Info=False;User ID=toyib;Data Source=csc"
    cstm.CursorLocation = adUseClient

End Sub

Public Sub ktgri()

    Set kategori = New ADODB.Recordset
    kategori.Open "SELECT company FROM customer order by company", cstm

End Sub
