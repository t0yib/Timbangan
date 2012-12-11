VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   2760
      TabIndex        =   15
      Top             =   2400
      Width           =   5175
      Begin VB.CommandButton cmbkeluar 
         Caption         =   "Keluar"
         Height          =   615
         Left            =   4200
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmbhapus 
         Caption         =   "Hapus"
         Height          =   615
         Left            =   3000
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmbsimpan 
         Caption         =   "Simpan"
         Height          =   615
         Left            =   1680
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmbbaru 
         Caption         =   "Baru"
         Height          =   615
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7815
      Begin VB.TextBox txtno 
         Height          =   405
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtnis 
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txttelpon 
         Height          =   375
         Left            =   4560
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtalamat 
         Height          =   1095
         Left            =   4560
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtnama 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "No Anggota"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "NIs"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Nama"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Alamat"
         Height          =   375
         Left            =   3600
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "No Telpon"
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Nama "
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pengurutan"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   2535
      Begin VB.OptionButton Option1 
         Caption         =   "No Anggota"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "KLIK 2 KALI UNTUk MERUBAH"
      Top             =   3480
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4471
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub fromkosong()
txtno.Text = ""
txtalamat.Text = ""
txtnama.Text = ""
txtnis.Text = ""
txttelpon.Text = ""

End Sub

Sub frommati()
txtno.Enabled = False
txtnama.Enabled = False
txtalamat.Enabled = False
txttelpon.Enabled = False
txtnis.Enabled = False

txtno.BackColor = &HFFC0C0
txtnama.BackColor = &HFFC0C0
txtalamat.BackColor = &HFFC0C0
txttelpon.BackColor = &HFFC0C0
txtnis.BackColor = &HFFC0C0

End Sub

Sub formhidup()
txtno.Enabled = True
txtnama.Enabled = True
txtalamat.Enabled = True
txttelpon.Enabled = True
txtnis.Enabled = True

txtno.BackColor = &HFFFFFF
txtnama.BackColor = &HFFFFFF
txtalamat.BackColor = &HFFFFFF
txttelpon.BackColor = &HFFFFFF
txtnis.BackColor = &HFFFFFF
End Sub

Sub fromnormal()
Call fromkosong
Call frommati

cmbbaru.Enabled = True
cmbhapus.Enabled = False
cmbsimpan.Enabled = False
cmbkeluar.Caption = "Keluar"
End Sub

Sub buatkode()
Dim kd As String

Set rs_anggota = New ADODB.Recordset
rs_anggota.Open "SELECT * FROM anggota order by no_anggota", _
dbkoneksi, adOpenDynamic, adLockBatchOptimistic

rs_anggota.Requery
With rs_anggota
If .BOF Then
Dim kodebaru As String
'jika tabel kosong kodenya
txtno.Text = "A0001"
kodebaru = "A0001"
Exit Sub


    Else
    'jika pelanggan lebih dari 1 baris
      'angka pada kode terakhir ditambah 1
      .MoveLast
      kd = !no_anggota
      kd = Val(Right(kd, 1))
      kd = kd + 1
      End If
      kodebaru = "A" + Format(kd, "0000")
      End With
      txtno.Text = kodebaru
      
      txtnis.SetFocus
      
End Sub

Private Sub cmbbaru_Click()
Call formhidup
Call buatkode

 AKSIDATA = "DATABARU"
    
    txtno.Enabled = False
    txtnis.SetFocus
    
    cmbbaru.Enabled = False
    cmbsimpan.Enabled = True
    cmbhapus.Enabled = False
    cmbkeluar.Caption = "normal"
End Sub

Private Sub cmbhapus_Click()
Tanya = MsgBox("YAKIN AKAN MENGHAPUS DATA INI?" _
        & vbCrLf & " no : " & txtno.Text + vbCrLf _
        & " nama : " & txtnama.Text + vbCrLf & "", _
         vbYesNo + vbQuestion, "Awass")

    If Tanya = vbYes Then
        SQL = "DELETE FROM anggota WHERE " _
        & " no_anggota='" & txtno.Text & "'"
        
        dbkoneksi.Execute SQL, , adCmdText
        
      
        Call fromnormal
        Call Form_Load
    Else
        Exit Sub
    End If
End Sub

Private Sub cmbkeluar_Click()
 If cmbkeluar.Caption = "Keluar" Then
        Unload Me
    Else
        Call fromnormal
    End If
End Sub

Private Sub cmbsimpan_Click()
If txtnis.Text = "" Then
     MsgBox "NIS TIDAK BOLEH KOSONG", _
     vbInformation + vbOKOnly, "ERROR"
         txtnis.SetFocus
         
ElseIf txtnama.Text = "" Then
 MsgBox "NAMA TIDAK BOLEH KOSONG", _
 vbInformation + vbOKOnly, "ERROR"
 txtnama.SetFocus
         
 
 
 ElseIf txtalamat.Text = "" Then
     MsgBox "ALAMAT TIDAK BOLEH KOSONG", _
     vbInformation + vbOKOnly, "ERROR"
         txtalamat.SetFocus
         

    Else
      If AKSIDATA = "DATABARU" Then
           SQLsimpan = ""
           SQLsimpan = "INSERT INTO anggota" _
              & "(no_anggota,nis,nama,alamat,notelpon)" _
              & "VALUES ('" & txtno.Text & "' , '" _
              & txtnis.Text & " ', '" _
              & txtnama.Text & " ', '" _
              & txtalamat.Text & "', '" _
              & txttelpon.Text & "' ) "
              dbkoneksi.Execute SQLsimpan, , adCmdText
              
              Call fromnormal
              Call Form_Load
               MsgBox "DATA TELAH DI SIMPAN", _
             vbInformation + vbOKOnly, "pemberitahuan"
             
            ElseIf AKSIDATA = "DATALAMA" Then
            SQLubah = ""
            SQLubah = "UPDATE anggota " _
                & " SET nis = '" & txtnis.Text & "' ,  " _
                 & " nama = '" & txtnama.Text & "' , " _
                & " alamat = '" & txtalamat.Text & "' , " _
                & " notelpon = ' " & txttelpon.Text & "' " _
                & " WHERE no_anggota = '" & txtno.Text & "'"
                dbkoneksi.Execute SQLubah, , adCmdText
              
              Call fromnormal
              Call Form_Load
               MsgBox "DATA TELAH DI UBAH", _
             vbInformation + vbOKOnly, "pemberitahuan"
               Else
            MsgBox "TIDAK ADA AKSI"
        End If
        
      
            
End If

End Sub

Private Sub DataGrid1_Click()
cmbhapus.Enabled = True
    cmbsimpan.Enabled = True
    cmbkeluar.Caption = "normal"
    cmbbaru.Enabled = False
    
     ' Status Ubah Data
    AKSIDATA = "DATALAMA"
    
    Call formhidup
    txtno.Enabled = False
   
    

Dim rs_anggota As New ADODB.Recordset
rs_anggota.Open "SELECT * FROM anggota WHERE no_anggota= '" & DataGrid1.Columns(0) & "'", dbkoneksi, adOpenDynamic, adLockBatchOptimistic
   txtno.Text = rs_anggota!no_anggota
        txtnis.Text = rs_anggota!nis
        txtnama.Text = rs_anggota!Nama
        txtalamat.Text = rs_anggota!alamat
        txttelpon.Text = rs_anggota!notelpon
End Sub
Private Sub Form_Load()
Call BukaDatabase
Dim rs_golongan As New ADODB.Recordset
rs_golongan.Open "SELECT * FROM anggota", dbkoneksi, adOpenDynamic, adLockBatchOptimistic
Set DataGrid1.DataSource = rs_golongan

Call fromnormal



End Sub

