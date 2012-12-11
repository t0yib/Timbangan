VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   16545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Add"
      Height          =   375
      Left            =   10800
      TabIndex        =   20
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      Height          =   375
      Left            =   12240
      TabIndex        =   19
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   375
      Left            =   13680
      TabIndex        =   18
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   15120
      TabIndex        =   17
      Top             =   4560
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9360
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "select custname,company,address1,phone1,fax,city,zipcode,country from customer"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FormCustomer.frx":0000
      Height          =   3015
      Left            =   5160
      TabIndex        =   16
      Top             =   1320
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
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
   Begin VB.TextBox Text8 
      DataField       =   "country"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   15
      Top             =   4680
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      DataField       =   "zipcode"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   14
      Top             =   4200
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      DataField       =   "city"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   13
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      DataField       =   "fax"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   12
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      DataField       =   "phone1"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   11
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      DataField       =   "address1"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   10
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      DataField       =   "company"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      DataField       =   "custname"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label8 
      Caption         =   "Country"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Zip Code"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "City"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Fax"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Phone"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Company"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "FormCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    FormCustomer.Enabled = False
    FormMain.Enabled = True
    FormMain.Show
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
Command4.Enabled = False
Command2.Enabled = False
Command1.Enabled = False
DataGrid1.Enabled = False

Else
Command3.Caption = "Edit"
Command4.Enabled = True
Command2.Enabled = True
Command1.Enabled = True
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
DataGrid1.Refresh
End If

End Sub

Private Sub Command4_Click()

If Command4.Caption = "Add" Then
Command4.Caption = "Save"
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Command3.Enabled = False
Command2.Enabled = False
Command1.Enabled = False
DataGrid1.Enabled = False
Adodc1.Recordset.AddNew

Else
Command4.Caption = "Add"
Command3.Enabled = True
Command2.Enabled = True
Command1.Enabled = True
DataGrid1.Enabled = True
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Then
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
DataGrid1.Refresh
End If

End Sub
