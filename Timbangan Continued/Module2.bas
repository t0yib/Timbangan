Attribute VB_Name = "Module2"
Public strcustomer As String
Public conn As New ADODB.Connection
Public rsUser As New ADODB.Recordset

Public Function Customer() As Boolean
    
'settingan koneksi
        On Error GoTo er
        
'koneksi string ke mysql konektor
        strcustomer = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=" & "localhost" & ";DATABASE=" & "csc" & ";UID=" & "root" & ";PWD=" & "" & ";PORT=" & "3306" & ";OPTION=3"

    If conn.State = adStateOpen Then conn.Close

    conn.Open strcustomer
    conn.CursorLocation = adUseClient

'buka tabel database
    rsUser.Open "SELECT * FROM customer", strcustomer, adOpenKeyset, adLockOptimistic

    If conn.State = adStateOpen Then
        Customer = True
        Exit Function
    Else
        Customer = False
        Exit Function
    End If
    
Exit Function

er:
    Customer = False
    MsgBox "Gagal Loading Database", vbInformation, "Database Error"
    
End Function
