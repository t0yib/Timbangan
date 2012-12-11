VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Timbang 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PT. Angkasa Pura Solusi - CSC Transaction System"
   ClientHeight    =   9015
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   15690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   15690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text17 
      Height          =   495
      Left            =   1440
      TabIndex        =   58
      Top             =   960
      Width           =   4335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Generate CSC"
      Height          =   495
      Left            =   120
      TabIndex        =   57
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      DataField       =   "country"
      DataSource      =   "Adodc6"
      Height          =   285
      Left            =   1800
      TabIndex        =   56
      Top             =   8640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text15 
      DataField       =   "city"
      DataSource      =   "Adodc6"
      Height          =   285
      Left            =   240
      TabIndex        =   55
      Top             =   8640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text14 
      DataField       =   "phone1"
      DataSource      =   "Adodc6"
      Height          =   285
      Left            =   4920
      TabIndex        =   54
      Top             =   8280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text13 
      DataField       =   "address1"
      DataSource      =   "Adodc6"
      Height          =   285
      Left            =   3360
      TabIndex        =   53
      Top             =   8280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text12 
      DataField       =   "company"
      DataSource      =   "Adodc6"
      Height          =   285
      Left            =   1800
      TabIndex        =   52
      Top             =   8280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      DataField       =   "custname"
      DataSource      =   "Adodc6"
      Height          =   285
      Left            =   240
      TabIndex        =   51
      Top             =   8280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   3120
      Top             =   7440
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   582
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
      RecordSource    =   "csc"
      Caption         =   "Adodc6"
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   375
      Left            =   3720
      Top             =   8640
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
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
      RecordSource    =   "comodity"
      Caption         =   "Adodc5"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   4680
      Top             =   7680
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      RecordSource    =   "select flightcompany from flight"
      Caption         =   "Adodc4"
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Form1.frx":0000
      DataField       =   "flightcompany"
      DataSource      =   "Adodc4"
      Height          =   315
      Left            =   6000
      TabIndex        =   48
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      ListField       =   "flightcompany"
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Form1.frx":0015
      DataField       =   "company"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   480
      TabIndex        =   47
      Top             =   2520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      ListField       =   "company"
      Text            =   "Pilih Customer"
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      DataField       =   "flightnum"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   1440
      TabIndex        =   46
      Top             =   7800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "SAVE"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9360
      TabIndex        =   45
      Top             =   8400
      Width           =   1815
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      DataField       =   "kg"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   360
      TabIndex        =   44
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10920
      TabIndex        =   43
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      DataField       =   "cscno1"
      DataSource      =   "Adodc3"
      Height          =   285
      Left            =   360
      TabIndex        =   42
      Top             =   7800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      DataField       =   "cscno1"
      DataSource      =   "Adodc6"
      Height          =   285
      Left            =   1440
      TabIndex        =   41
      Top             =   7440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      TabIndex        =   40
      Top             =   7080
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   4680
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      RecordSource    =   "select cscno1 from csc"
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3120
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
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
      RecordSource    =   "csccom"
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3120
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
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
   Begin VB.CommandButton Command1 
      Caption         =   "RUN"
      Height          =   375
      Left            =   11280
      TabIndex        =   26
      Top             =   8400
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000003&
      DataField       =   "cscno1"
      DataSource      =   "Adodc6"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "@Adobe Heiti Std R"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   21
      Top             =   1560
      Width           =   4335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "create new"
      BeginProperty Font 
         Name            =   "Bell Gothic Std Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox display 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   7080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3960
      Top             =   6360
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      DataField       =   "spclcode"
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   8040
      TabIndex        =   14
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      DataField       =   "pcs"
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   8040
      TabIndex        =   10
      Top             =   2880
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10920
      TabIndex        =   9
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   8040
      TabIndex        =   7
      Text            =   " . . . "
      Top             =   1560
      Width           =   6975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "[ + ] Add / Edit Customer"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   13200
      TabIndex        =   0
      Top             =   8400
      Width           =   1815
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4560
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   5535
      Begin VB.Frame Frame2 
         BackColor       =   &H80000005&
         Caption         =   "Consigner Data View"
         Height          =   2535
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   5055
         Begin VB.Label Label19 
            BackColor       =   &H8000000E&
            Caption         =   "country"
            DataField       =   "country"
            DataSource      =   "Adodc1"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   2040
            Width           =   3615
         End
         Begin VB.Label Label18 
            BackColor       =   &H8000000E&
            Caption         =   "zipcode"
            DataField       =   "zipcode"
            DataSource      =   "Adodc1"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   1800
            Width           =   3615
         End
         Begin VB.Label Label17 
            BackColor       =   &H8000000E&
            Caption         =   "city"
            DataField       =   "city"
            DataSource      =   "Adodc1"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   1560
            Width           =   3615
         End
         Begin VB.Label Label16 
            BackColor       =   &H8000000E&
            Caption         =   "fax"
            DataField       =   "fax"
            DataSource      =   "Adodc1"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   1320
            Width           =   3615
         End
         Begin VB.Label Label15 
            BackColor       =   &H8000000E&
            Caption         =   "telp"
            DataField       =   "phone1"
            DataSource      =   "Adodc1"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   1080
            Width           =   3615
         End
         Begin VB.Label Label14 
            BackColor       =   &H8000000E&
            Caption         =   "alamat1"
            DataField       =   "address1"
            DataSource      =   "Adodc1"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   840
            Width           =   3615
         End
         Begin VB.Label Label13 
            BackColor       =   &H8000000E&
            Caption         =   "pt"
            DataField       =   "company"
            DataSource      =   "Adodc1"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label12 
            BackColor       =   &H8000000E&
            Caption         =   "nama"
            DataField       =   "custname"
            DataSource      =   "Adodc1"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5400
         Y1              =   960
         Y2              =   960
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Commodity Overview"
      Height          =   6135
      Left            =   6000
      TabIndex        =   11
      Top             =   2040
      Width           =   9135
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Form1.frx":002A
         Height          =   975
         Left            =   6840
         TabIndex        =   50
         Top             =   5040
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1720
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
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Form1.frx":003F
         DataField       =   "comname"
         DataSource      =   "Adodc5"
         Height          =   315
         Left            =   2040
         TabIndex        =   49
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         ListField       =   "comname"
         Text            =   ""
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form1.frx":0054
         Height          =   1095
         Left            =   6840
         TabIndex        =   29
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1931
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
      Begin VB.CommandButton Command4 
         Caption         =   "Edit Commodity"
         Height          =   375
         Left            =   7560
         TabIndex        =   27
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Add "
         Enabled         =   0   'False
         Height          =   375
         Left            =   4920
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         DataField       =   "awb"
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   16
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label21 
         Caption         =   "TTL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7680
         TabIndex        =   39
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "Weight Total  :"
         Height          =   255
         Left            =   6360
         TabIndex        =   38
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "AWB / SMU Code"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "SPCL Code"
         Height          =   255
         Left            =   840
         TabIndex        =   15
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantity"
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Commodity"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   4200
      TabIndex        =   28
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   24
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Shipper :"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   23
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kg"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7440
      TabIndex        =   22
      Top             =   120
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   7320
      Top             =   120
      Width           =   735
   End
   Begin VB.Label LabelDate 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   " dd/MM/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   2040
      TabIndex        =   19
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label LabelTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "HH:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "Remarks"
      Height          =   255
      Left            =   8040
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Flight No / Info"
      Height          =   255
      Left            =   6000
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Consignment Security Certificate"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   615
      Left            =   9000
      TabIndex        =   1
      Top             =   120
      Width           =   6855
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   8040
      Top             =   120
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   8040
      Shape           =   2  'Oval
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape4 
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   7095
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "TImbang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public statusadmin As Boolean

Private Sub Command3_Click()
    Customer.statusadmin = True
    Customer.Show
    TImbang.Enabled = False
    
End Sub

Private Sub Command4_Click()
    Comodity.statusadmin = True
    Comodity.Show
    TImbang.Enabled = False
    
End Sub


Private Sub Command5_Click()
    Dim a As String
    Dim b As Integer
    
    Command7.Enabled = True
    Command9.Enabled = True
    Command8.Enabled = True
    DataGrid1.Enabled = True
        If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or DataCombo3.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text10.Text = "Pilih Flight" Then
        Adodc2.Recordset.Cancel
        Adodc2.Refresh
        Else
        Adodc2.Recordset.Update
        End If
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    DataCombo3.Enabled = False
    DataGrid1.Refresh
    
    Text9.Text = display.Text
    b = CInt(a)
    b = b + b
    'Text9.Text = a
    
End Sub

Private Sub Command6_Click()
    
    Text6.Enabled = True
    Text6.BackColor = &H80000005
    DataCombo1.Enabled = True
    DataCombo2.Enabled = True
    Command7.Enabled = True
    Command2.Enabled = False
    Command8.Enabled = True
    Command9.Enabled = True
    DataGrid2.Enabled = False
    Adodc6.Recordset.AddNew
    
    Text11.Text = Label12.Caption
    Text12.Text = Label13.Caption
    Text13.Text = Label14.Caption
    Text14.Text = Label15.Caption
    Text15.Text = Label17.Caption
    Text16.Text = Label19.Caption
    
    
    Text6.Text = "CSC." + Text5.Text
    'Text10.Text = DataCombo2.Text
   

End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    Command9.Enabled = False
    Command8.Enabled = False
    DataCombo3.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    DataGrid1.Enabled = False
    Adodc2.Recordset.AddNew
    
    
    
    
    

End Sub

Private Sub Command8_Click()

    Command2.Enabled = True
    Adodc6.Recordset.Update
    DataGrid2.Enabled = True
    
End Sub



Private Sub Form_Load()
Dim i As Byte
Dim j As Byte
Dim k As Byte
Dim l As Double
Dim aa As Integer
Dim data As String
Dim buffer As String
Dim char As String
Dim x As String
Dim stab As Integer
Dim dt As Integer
Dim mt As Integer
Dim yr As Integer
Dim hr As Integer
Dim mn As Integer
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
Dim ff As String
Dim m As String
Dim n As Integer

l = 0.5

Label10.Caption = Login.Text1

dt = Day(Now)
b = Str(dt)
mt = Month(Now)
c = Str(mt)
yr = Year(Now)
d = Str(yr)
hr = Hour(Now)
e = Str(hr)
mn = Minute(Now)
f = Str(mn)
n = Second(Now)
n = Str(n)

Text5.Text = b + c + d + "." + e + f + n



End Sub

Private Sub Command1_Click()

    Command1.Enabled = False

    MSComm1.PortOpen = True
        Do
        DoEvents
            data = MSComm1.Input
            buffer = buffer + data
            i = Len(buffer)
            For j = 1 To i
                char = Mid(buffer, j, 1)
                    If char = "S" Then
                        For k = 1 To i
                        char = Mid(buffer, k, 1)
                            If char = Chr(13) Then
                            data = ""
                                stab = CInt(j)
                                    If stab = 1 Then
                                    Command5.Enabled = True
                                    Shape7.Visible = True
                                    Shape1.Visible = False
                                    Else
                                    Command5.Enabled = False
                                    Shape7.Visible = False
                                    Shape1.Visible = True
                                    End If
                                For i = j To k
                                char = Mid(buffer, i, 1)
                                    If (char = "1") Or (char = "2") Or (char = "3") Or (char = "4") Or (char = "5") Or (char = "6") Or (char = "7") Or (char = "8") Or (char = "9") Or (char = "0") Or (char = "-") Then
                                    data = data + char
                                    End If
                                i = i + l
                                l = l * 2
                                Next i
                                display.Text = data
                                Label11.Caption = data
                                buffer = ""
                                Exit For
                            End If
                        k = k + 1
                        Next k
                    End If
                    j = j + 1
            Next j
        Loop
        
        
End Sub

Private Sub Command2_Click()

    Result = MsgBox("Keluar Window CSC?", vbYesNo)
    If Result = vbYes Then
    End
    
    End If
    
End Sub

Private Sub Timer1_Timer()

    LabelTime.Caption = Time
    LabelDate.Caption = Date
    
End Sub
