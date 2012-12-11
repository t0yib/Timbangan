Attribute VB_Name = "Module1"
Option Explicit
Public dbkoneksi As New ADODB.Connection
Public rs_anggota As ADODB.Recordset



Public db_name As String
Public db_server As String
Public db_port As String
Public db_user As String
Public db_pass As String

Public StrKonekDb As String
Public strSQL, SQL As String
Public SQLubah, SQLsimpan, SQLhapus As String
Public Tanya As String
Public Status As String
Public AKSIDATA As String
Public UserId, NamaId As String

Public Sub BukaDatabase()
 On Error GoTo ok
    Set dbkoneksi = New ADODB.Connection
    db_name = "perpus"
    db_server = "localhost" 'ganti jika server anda ada di komputer lain
    db_port = "3306"    'default port is 3306
    db_user = "root"    'sebaiknya pakai username lain.
    db_pass = ""
  
StrKonekDb = "DRIVER={MySQL ODBC 3.51 Driver};" _
        & " SERVER=" & db_server & ";" _
        & "DATABASE=" & db_name & ";" _
        & "UID=" & db_user & ";" _
        & "PWD=" & db_pass & ";" _
       & "PORT=" & db_port & ";" _
       & "OPTION=3"

dbkoneksi.CursorLocation = adUseClient
    If dbkoneksi.State = adStateOpen Then
         dbkoneksi.Close
        Set dbkoneksi = New ADODB.Connection
        dbkoneksi.Open StrKonekDb
    Else
     
           dbkoneksi.Open StrKonekDb
            
    End If
    Exit Sub
ok:
    
    MsgBox "Koneksi ke server tidak berhasil !!!", 16, "ERROR"
   End
    
End Sub



