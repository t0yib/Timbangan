Attribute VB_Name = "ModuleComodity"
Option Explicit
Public com As New ADODB.Connection
Public comod As New ADODB.Recordset

Public Sub konekcomodity()

    Set com = New ADODB.Connection
    com.Open "Provider=MSDASQL.1;Persist Security Info=False;User ID=toyib;Data Source=csc"
    com.CursorLocation = adUseClient
    
End Sub

Public Sub cmdy()

Set comod = New ADODB.Recordset
comod.Open "SELECT comname FROM comodity ordeer by comname", com

End Sub

