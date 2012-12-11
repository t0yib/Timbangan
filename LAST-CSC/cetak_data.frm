VERSION 5.00
Begin VB.Form cetak_data 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   7080
      TabIndex        =   1
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9360
      TabIndex        =   0
      Top             =   6240
      Width           =   1815
   End
End
Attribute VB_Name = "cetak_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'If Command1.Caption = "PRINT" Then
'cetak
'Command1.Caption = "EXIT"

'Else
'Command1.Caption = "PRINT"
'FormMain.Show
'FormMain.Enabled = True
'Unload Me
'End If
 
End Sub

Sub cetak()
Printer.Font = "Arial"
Show
     CurrentX = 0
     CurrentY = 0
     Printer.FontSize = 18
     Printer.Print Tab(15); "CONSIGNMENT SECURITY CERTIFICATE";
     Printer.Font = "Courier New"
     Printer.FontSize = 10
     Printer.Print Tab(2); ""
     Printer.Print Tab(32); Format(Date, "dd/mm/yyyy"); " CSC No :"; FormMain.Text1.Text;
     Printer.Print Tab(2); "==============================================================================================";
     Printer.Print Tab(3); "                          CONSIGNOR NAME   :"; FormMain.Text2.Text;
     Printer.Print Tab(3); "                          COMPANY          :"; FormMain.Text3.Text;
     Printer.Print Tab(3); "                          ADDRESS          :"; FormMain.Text4.Text;
     Printer.Print Tab(3); "                          PHONE/FAX        :"; FormMain.Text5.Text;
     Printer.Print Tab(2); "==============================================================================================";
     Printer.FontSize = 9.5
     Printer.Print Tab(2); " COMMODITY | QTY  |     WEIGHT     |   DST CODE  |     AWB / SMU  "
     Printer.Print Tab(2); "==============================================================================================";
     
     Do While Not FormMain.Adodc3.Recordset.EOF
     Printer.Print Tab(2); FormMain.DataGrid1.Columns(0).Text & Space(1), FormMain.DataGrid1.Columns(1).Text & Space(1), FormMain.DataGrid1.Columns(2).Text & Space(1), FormMain.DataGrid1.Columns(3).Text & Space(1), FormMain.DataGrid1.Columns(4).Text & Space(1);
     FormMain.Adodc3.Recordset.MoveNext
     Loop
     
    Printer.FontSize = 9.5
    Printer.Print Tab(2); "==============================================================================================";
    Printer.Font = "Courier New"
    Printer.Print Tab(2); "==============================================================================================";
    Printer.Print Tab(5); "Remarks  : "; FormMain.Text47.Text;
    Printer.Print Tab(5); ""
    Printer.Print Tab(2); "    Shipper"; "               Regulated Agent Sign";
    Printer.Print Tab(2); "                                              ";
    Printer.Print Tab(2); "                                              ";
    Printer.Print Tab(2); "  (         )              (                ) ";
    Printer.EndDoc
    
    FormMain.Adodc3.Refresh
End Sub

Private Sub Command2_Click()
    FormMain.Show
    FormMain.Enabled = True
    Unload Me
End Sub

Private Sub Form_Load()

Font = "Arial"
Show
     CurrentX = 0
     CurrentY = 0
     FontSize = 18
     Print Tab(15); "CONSIGNMENT SECURITY CERTIFICATE";
     Font = "Courier New"
     FontSize = 10
     Print Tab(2); ""
     Print Tab(32); Format(Date, "dd/mm/yyyy"); " CSC No :"; FormMain.Text1.Text;
     Print Tab(2); "==============================================================================================";
     Print Tab(3); "                          CONSIGNOR NAME   :"; FormMain.Text2.Text;
     Print Tab(3); "                          COMPANY          :"; FormMain.Text3.Text;
     Print Tab(3); "                          ADDRESS          :"; FormMain.Text4.Text;
     Print Tab(3); "                          PHONE/FAX        :"; FormMain.Text5.Text;
     Print Tab(2); "==============================================================================================";
     FontSize = 9.5
     Print Tab(2); " COMMODITY | QTY  |     WEIGHT     |   DST CODE  |     AWB / SMU  "
     Print Tab(2); "==============================================================================================";
     
     FormMain.Adodc3.Refresh
     Do While Not FormMain.Adodc3.Recordset.EOF
     Print Tab(2); FormMain.DataGrid1.Columns(0).Text & Space(1), FormMain.DataGrid1.Columns(1).Text & Space(1), FormMain.DataGrid1.Columns(2).Text & Space(1), FormMain.DataGrid1.Columns(3).Text & Space(1), FormMain.DataGrid1.Columns(4).Text & Space(1);
     FormMain.Adodc3.Recordset.MoveNext
     Loop
     
FontSize = 9.5
    Print Tab(2); "==============================================================================================";
Font = "Courier New"
    Print Tab(2); "==============================================================================================";
    Print Tab(5); "Remarks  : "; FormMain.Text47.Text;
    Print Tab(5); ""
    Print Tab(2); "    Shipper"; "               Regulated Agent Sign";
    Print Tab(2); "                                              ";
    Print Tab(2); "                                              ";
    Print Tab(2); "  (         )              (                ) ";
    
    
    FormMain.Adodc3.Refresh

End Sub
