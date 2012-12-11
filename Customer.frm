VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Customer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PT. Angkasa Pura Solusi - CSC Transaction System - Customer Edit"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   13110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Add New"
      Height          =   375
      Left            =   10800
      TabIndex        =   26
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      Height          =   375
      Left            =   10800
      TabIndex        =   25
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   3840
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Customer.frx":0000
      Height          =   1815
      Left            =   240
      TabIndex        =   22
      Top             =   4560
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   0   'False
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   240
      Top             =   6720
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;User ID=toyib;Data Source=radb"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;User ID=toyib;Data Source=radb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "customer"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   375
      Left            =   10800
      TabIndex        =   11
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text10 
      DataField       =   "custdesc"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   975
      Left            =   7320
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text9 
      DataField       =   "zipcode"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   9
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox Text8 
      DataField       =   "country"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      DataField       =   "city"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      DataField       =   "fax"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      DataField       =   "phone1"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      DataField       =   "address2"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   855
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      DataField       =   "address1"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   855
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      DataField       =   "company"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      DataField       =   "custname"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   10920
      TabIndex        =   0
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "NPWP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   24
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Cust. Desc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   21
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Zip Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   20
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Country"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   19
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   18
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   17
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Alamat 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Alamat 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Perusahaan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nama"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public statusadmin As Boolean
 
Private Sub Command1_Click()

    TImbang.statusadmin = True
    TImbang.Show
    Customer.Enabled = False
    TImbang.Enabled = True
    Unload Me

End Sub

Private Sub Command2_Click()
Hapus = MsgBox("Yakin ingin menghapus ?", vbQuestion + vbYesNo, "Perhatian")
If Hapus = vbYes Then
    Adodc1.Recordset.Delete
    DataGrid1.Refresh
End If
   
End Sub

Private Sub Command3_Click()
If Command3.Caption = "Edit" Then
Command3.Caption = "Save"
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Command4.Enabled = False
Command2.Enabled = False
DataGrid1.Enabled = False

Else
Command3.Caption = "Edit"
Command4.Enabled = True
Command2.Enabled = True
DataGrid1.Enabled = True
Adodc1.Recordset.Update
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
DataGrid1.Refresh
End If

End Sub

Private Sub Command4_Click()

If Command4.Caption = "Add New" Then
Command4.Caption = "Save"
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Command3.Enabled = False
Command2.Enabled = False
DataGrid1.Enabled = False
Adodc1.Recordset.AddNew
Text1.SetFocus
Text2.SetFocus
Else
Command4.Caption = "Add New"
Command3.Enabled = True
Command2.Enabled = True
DataGrid1.Enabled = True
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text5.Text = "" Or Text7.Text = "" Or Text8.Text = "" Then
        Adodc1.Recordset.Cancel
        Adodc1.Refresh
    Else
        Adodc1.Recordset.Update
    End If
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
DataGrid1.Refresh

End If

End Sub

