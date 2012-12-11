Attribute VB_Name = "ModuleLogin"
Public strkoneksi As String
Public conn As New ADODB.Connection
Public rsUser As New ADODB.Recordset

Public Function Koneksi() As Boolean
    
'settingan koneksi
        On Error GoTo er
        
'koneksi string ke mysql konektor
        strkoneksi = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=" & "localhost" & ";DATABASE=" & "csc" & ";UID=" & "root" & ";PWD=" & "" & ";PORT=" & "3306" & ";OPTION=3"

    If conn.State = adStateOpen Then conn.Close

    conn.Open strkoneksi
    conn.CursorLocation = adUseClient

'buka tabel database
    rsUser.Open "SELECT username,userpass,lastname FROM user", strkoneksi, adOpenKeyset, adLockOptimistic

    If conn.State = adStateOpen Then
        Koneksi = True
        Exit Function
    Else
        Koneksi = False
        Exit Function
    End If
    
Exit Function

er:
    Koneksi = False
    MsgBox "Gagal Loading Database", vbInformation, "Database Error"
    
End Function


Public Function Customer() As Boolean
    
'settingan koneksi
        On Error GoTo er
        
'koneksi string ke mysql konektor
        strcustomer = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=" & "192.168.1.112" & ";DATABASE=" & "csc" & ";UID=" & "toyib" & ";PWD=" & "" & ";PORT=" & "3306" & ";OPTION=3"

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
    
er:
    Customer = False
    MsgBox "Gagal Loading Database", vbInformation, "Database Error"
    
End Function




