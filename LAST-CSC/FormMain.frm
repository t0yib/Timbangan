VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   16845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   16845
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text59 
      DataField       =   "npwp"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   6120
      TabIndex        =   95
      Text            =   "Text58"
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text58 
      DataField       =   "custid"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   5760
      TabIndex        =   94
      Text            =   "Text58"
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text57 
      DataField       =   "npwp"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   5040
      TabIndex        =   93
      Text            =   "Text57"
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text10 
      DataField       =   "custid"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4680
      TabIndex        =   92
      Text            =   "Text10"
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command12 
      Caption         =   "PRINT"
      Height          =   495
      Left            =   9960
      TabIndex        =   90
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8880
      Top             =   600
   End
   Begin VB.CommandButton Command11 
      Caption         =   "RUN"
      Height          =   495
      Left            =   11400
      TabIndex        =   87
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox Text56 
      DataField       =   "cscid"
      DataSource      =   "Adodc4"
      Height          =   375
      Left            =   10800
      TabIndex        =   86
      Text            =   "Text56"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text55 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12240
      TabIndex        =   83
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text54 
      DataField       =   "awb"
      DataSource      =   "Adodc5"
      Height          =   405
      Left            =   3960
      TabIndex        =   82
      Text            =   "Text17"
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text53 
      DataField       =   "spclcode"
      DataSource      =   "Adodc5"
      Height          =   405
      Left            =   3600
      TabIndex        =   81
      Text            =   "Text17"
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text52 
      DataField       =   "flightnum"
      DataSource      =   "Adodc5"
      Height          =   405
      Left            =   3240
      TabIndex        =   80
      Text            =   "Text17"
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text51 
      DataField       =   "kg"
      DataSource      =   "Adodc5"
      Height          =   405
      Left            =   2880
      TabIndex        =   79
      Text            =   "Text17"
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text50 
      DataField       =   "pcs"
      DataSource      =   "Adodc5"
      Height          =   405
      Left            =   2520
      TabIndex        =   78
      Text            =   "Text17"
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text49 
      DataField       =   "comname"
      DataSource      =   "Adodc5"
      Height          =   405
      Left            =   2160
      TabIndex        =   77
      Text            =   "Text17"
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text48 
      DataField       =   "cscno1"
      DataSource      =   "Adodc5"
      Height          =   405
      Left            =   1800
      TabIndex        =   76
      Text            =   "Text17"
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   375
      Left            =   480
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      UserName        =   "toyib"
      Password        =   ""
      RecordSource    =   "select cscid,cscno1,comname,pcs,kg,flightnum,spclcode,awb from csccom"
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
   Begin VB.TextBox Text47 
      Height          =   285
      Left            =   11520
      TabIndex        =   21
      Top             =   2760
      Width           =   5055
   End
   Begin VB.TextBox Text46 
      Height          =   285
      Left            =   15000
      TabIndex        =   20
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text45 
      Height          =   285
      Left            =   15000
      TabIndex        =   19
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text44 
      Height          =   285
      Left            =   11520
      TabIndex        =   18
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text43 
      Height          =   285
      Left            =   11520
      TabIndex        =   17
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text42 
      DataField       =   "cscid"
      DataSource      =   "Adodc5"
      Height          =   405
      Left            =   4320
      TabIndex        =   70
      Text            =   "Text17"
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text41 
      DataField       =   "processstat"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   10440
      TabIndex        =   69
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text40 
      DataField       =   "line"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   10080
      TabIndex        =   68
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text39 
      DataField       =   "shift"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   9720
      TabIndex        =   67
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text38 
      DataField       =   "username"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   9360
      TabIndex        =   66
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text37 
      DataField       =   "remarks"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   9000
      TabIndex        =   65
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text36 
      DataField       =   "passangerid"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   8640
      TabIndex        =   64
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text35 
      DataField       =   "passanger"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   8280
      TabIndex        =   63
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text34 
      DataField       =   "driverid"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   7920
      TabIndex        =   62
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text33 
      DataField       =   "driver"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   7560
      TabIndex        =   61
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text32 
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   7200
      TabIndex        =   60
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text31 
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   6840
      TabIndex        =   59
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text30 
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   6480
      TabIndex        =   58
      Text            =   "Text"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text29 
      DataField       =   "totalkg"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   6120
      TabIndex        =   57
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text28 
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   5760
      TabIndex        =   56
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text27 
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   5400
      TabIndex        =   55
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text26 
      DataField       =   "zipcode"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   5040
      TabIndex        =   54
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text25 
      DataField       =   "country"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   4680
      TabIndex        =   53
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text24 
      DataField       =   "city"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   4320
      TabIndex        =   52
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text23 
      DataField       =   "fax"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   3960
      TabIndex        =   51
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text22 
      DataField       =   "phone1"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   3600
      TabIndex        =   50
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text21 
      DataField       =   "address1"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   3240
      TabIndex        =   49
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text20 
      DataField       =   "company"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   2880
      TabIndex        =   48
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text19 
      DataField       =   "custname"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   2520
      TabIndex        =   47
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text18 
      DataField       =   "cscdate"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   2160
      TabIndex        =   46
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text17 
      DataField       =   "cscno1"
      DataSource      =   "Adodc4"
      Height          =   405
      Left            =   1800
      TabIndex        =   45
      Text            =   "Text17"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   375
      Left            =   480
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      UserName        =   "toyib"
      Password        =   ""
      RecordSource    =   "csc"
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
   Begin VB.CommandButton Command7 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   14880
      TabIndex        =   38
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "SAVE"
      Enabled         =   0   'False
      Height          =   495
      Left            =   12960
      TabIndex        =   37
      Top             =   6600
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   3120
      Top             =   6120
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
      UserName        =   "toyib"
      Password        =   ""
      RecordSource    =   "temporary"
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
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12240
      TabIndex        =   33
      Top             =   3240
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Caption         =   "COMMODITY DATA VIEW"
      Enabled         =   0   'False
      Height          =   1935
      Left            =   5280
      TabIndex        =   24
      Top             =   1920
      Width           =   4695
      Begin VB.CommandButton Command8 
         Caption         =   "SAVE WEIGHT"
         Height          =   495
         Left            =   3120
         TabIndex        =   36
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text14 
         DataField       =   "spclcode"
         DataSource      =   "Adodc3"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text15 
         DataField       =   "awbsmu"
         DataSource      =   "Adodc3"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         DataField       =   "comquantity"
         DataSource      =   "Adodc3"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         DataField       =   "comweight"
         DataSource      =   "Adodc3"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   28
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text11 
         DataField       =   "comname"
         DataSource      =   "Adodc3"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Kg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   31
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "AWB /  SMU  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "DEST. CODE  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Weight  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Quantity  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Commodity  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "COMMODITY"
      Height          =   2175
      Left            =   5280
      TabIndex        =   23
      Top             =   4080
      Width           =   9975
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FormMain.frx":0000
         Height          =   1695
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   2990
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
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
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "FormMain.frx":0015
      Height          =   315
      Left            =   5280
      TabIndex        =   12
      Top             =   1440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "flightcompany"
      Text            =   "Flight"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   6120
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
      UserName        =   "toyib"
      Password        =   ""
      RecordSource    =   "select custid,custname,company,address1,phone1,fax,city,zipcode,country,npwp from customer"
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "FormMain.frx":002A
      DataField       =   "company"
      Height          =   315
      Left            =   600
      TabIndex        =   2
      Top             =   2520
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "company"
      BoundColumn     =   ""
      Text            =   "Customer"
   End
   Begin VB.Frame Frame1 
      Caption         =   "CUSTOMER OVERVIEW"
      Height          =   3855
      Left            =   480
      TabIndex        =   22
      Top             =   2160
      Width           =   4455
      Begin VB.TextBox Text7 
         DataField       =   "city"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "kota"
         Top             =   2640
         Width           =   4215
      End
      Begin VB.TextBox Text6 
         DataField       =   "fax"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "fax"
         Top             =   2280
         Width           =   4215
      End
      Begin VB.TextBox Text5 
         DataField       =   "phone1"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "telp"
         Top             =   1920
         Width           =   4215
      End
      Begin VB.TextBox Text4 
         DataField       =   "address1"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Text            =   "alamat"
         Top             =   1560
         Width           =   4215
      End
      Begin VB.TextBox Text3 
         DataField       =   "company"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "perush"
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox Text9 
         DataField       =   "country"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Text            =   "country"
         Top             =   3360
         Width           =   4215
      End
      Begin VB.TextBox Text8 
         DataField       =   "zipcode"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "zipcode"
         Top             =   3000
         Width           =   4215
      End
      Begin VB.TextBox Text2 
         DataField       =   "custname"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "nama"
         Top             =   840
         Width           =   4215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Edit Customer"
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Text            =   "text"
      Top             =   1440
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New CSC"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   975
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4680
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1800
      Top             =   6120
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
      UserName        =   "toyib"
      Password        =   ""
      RecordSource    =   "select flightcompany from flight"
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
   Begin VB.Frame Frame4 
      Enabled         =   0   'False
      Height          =   2655
      Left            =   15360
      TabIndex        =   39
      Top             =   3600
      Width           =   1335
      Begin VB.CommandButton Command6 
         Caption         =   "DELETE"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "DOWN"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "UP"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "SAVE"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "ADD"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   5640
      TabIndex        =   91
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "Label16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      TabIndex        =   89
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Label16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   88
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label15 
      Caption         =   "pcs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13320
      TabIndex        =   85
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "Total Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   84
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label13 
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10200
      TabIndex        =   75
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Pasgr ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13560
      TabIndex        =   74
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Pasgr. Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13560
      TabIndex        =   73
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Driver ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10200
      TabIndex        =   72
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Driver Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10200
      TabIndex        =   71
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Kg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13320
      TabIndex        =   34
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Total Weight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   32
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   240
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   7920
      Top             =   120
      Width           =   855
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   5520
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Dim b As String
Dim c As Integer
Dim d As String
Dim e As String
Dim f As Integer
Dim m As Integer
Dim n As Integer
Dim o As Integer
Dim p As Integer
Dim q As Integer
Dim r As Integer
Dim dy As String
Dim mt As String
Dim yr As String
Dim hr As String
Dim mnt As String
Dim sc As String
Dim s As String
Dim t As Integer
Dim u As Integer
Dim id As String
Dim npwp As String
Dim aa As Integer

    

Private Sub Command1_Click()
    
    Command7.Enabled = False
    Command1.Enabled = False
    Command5.Enabled = True
    Command4.Enabled = True
    Command9.Enabled = True
    Command6.Enabled = True
    Frame4.Enabled = True
    Adodc4.Recordset.AddNew
    
    m = Day(Now)
    dy = Str(m)
    n = Month(Now)
    mt = Str(n)
    o = Year(Now)
    yr = Str(o)
    p = Hour(Now)
    hr = Str(p)
    q = Minute(Now)
    mnt = Str(q)
    r = Second(Now)
    sc = Str(r)
    
    
    s = "CSC." + dy + mt + yr + "." + hr + mt + sc
    Text1.Text = Replace(s, " ", "")
    
    id = Text10.Text
    npwp = Text57.Text
    

End Sub

Private Sub Command10_Click()

    Text17.Text = Text1.Text
    Text19.Text = Text2.Text
    Text20.Text = Text3.Text
    Text21.Text = Text4.Text
    Text22.Text = Text5.Text
    Text23.Text = Text6.Text
    Text24.Text = Text7.Text
    Text25.Text = Text9.Text
    Text26.Text = Text8.Text
    'Text27.Text = a
    'Text28.Text = b
    Text29.Text = Text16.Text
    'Text30.Text = DataCombo2.Text
    'Text31.Text = d
    'Text32.Text = e
    Text33.Text = Text43.Text
    Text34.Text = Text44.Text
    Text35.Text = Text45.Text
    Text36.Text = Text46.Text
    Text37.Text = Text47.Text
    Text41.Text = 1
    Adodc4.Recordset.Update
    
    Command7.Enabled = True
    Command10.Enabled = False
    Command1.Enabled = True
    Frame4.Enabled = False
    
    
        With Adodc3.Recordset
        Do Until .EOF
            .Delete
            .MoveNext
        Loop
        End With
        
    Adodc3.Refresh
    DataGrid1.Refresh
    Text16.Text = ""
    
    Text55.Text = ""
    Text1.Text = ""
    Text58.Text = id
    Text59.Text = npwp
    
    Command5.Enabled = False
    Command4.Enabled = False
    Command9.Enabled = False
    Command6.Enabled = False

    
    
End Sub

Private Sub Command11_Click()

        Command11.Enabled = False

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
                                    Command8.Enabled = True
                                    Shape1.FillColor = &HFF00&
                                    Else
                                    Command8.Enabled = False
                                    Shape1.FillColor = &HFF&
                                    End If
                                    For i = j To k
                                    char = Mid(buffer, i, 1)
                                    If (char = "1") Or (char = "2") Or (char = "3") Or (char = "4") Or (char = "5") Or (char = "6") Or (char = "7") Or (char = "8") Or (char = "9") Or (char = "0") Or (char = "-") Then
                                    data = data + char
                                    End If
                                    i = i + l
                                    l = l * 2
                                    Next i
                                    'display.Text = data
                                    Label18.Caption = data
                                    aa = data
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

Private Sub Command12_Click()

cetak_data.Show
Command10.Enabled = True
FormMain.Enabled = False

End Sub

Private Sub Command2_Click()

    FormMain.Enabled = False
    FormCustomer.Enabled = True
    FormCustomer.Show
    Unload Me
    
End Sub


Private Sub Command3_Click()
    
    If Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or Text14.Text = "" Or Text15.Text = "" Then
        Adodc3.Recordset.Cancel
        Adodc5.Recordset.Cancel
        Adodc3.Refresh
        Adodc5.Refresh
    Else
        a = Text11.Text
        b = Text12.Text
        d = Text14.Text
        e = Text15.Text
        c = c + CInt(Text13.Text)
        Text16.Text = c
        Text48.Text = Text1.Text
        Text49.Text = a
        Text50.Text = b
        Text51.Text = Text13.Text
        Text52.Text = DataCombo2.Text
        Text53.Text = d
        Text54.Text = e
        Text55.Text = b
        
        Adodc3.Recordset.Update
        Adodc5.Recordset.Update
        
    End If
    Command3.Enabled = False
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command9.Enabled = True
    
    Text11.Enabled = False
    Text12.Enabled = False
    Text14.Enabled = False
    Text15.Enabled = False
    Frame3.Enabled = False
    f = 0
    DataGrid1.Enabled = True
    DataGrid1.Refresh

End Sub

Private Sub Command4_Click()

Adodc3.Recordset.MovePrevious

End Sub

Private Sub Command5_Click()
    
    Text42.Text = u
        
    Command4.Enabled = False
    Command6.Enabled = False
    Command9.Enabled = False
    DataGrid1.Enabled = False
    Text11.Enabled = True
    Text12.Enabled = True
    Text14.Enabled = True
    Text15.Enabled = True
    Frame3.Enabled = True
    Adodc3.Recordset.AddNew
    Adodc5.Recordset.AddNew
    Command5.Enabled = False
    

End Sub

Private Sub Command6_Click()

    Hapus = MsgBox("Yakin ingin menghapus ?", vbQuestion + vbYesNo, "Perhatian")
    If Hapus = vbYes Then
    Adodc3.Recordset.Delete
    DataGrid1.Refresh
    End If

End Sub

Private Sub Command7_Click()

    End

End Sub

Private Sub Command8_Click()

    t = t + 1
        
    Tambah = MsgBox("Ingin save ?", vbQuestion + vbYesNo, "Perhatian")
    If Tambah = vbYes Then
        f = f + (aa - 18)
        Command3.Enabled = True
    Else
        If t = 1 Then
        f = f + (aa - 18)
        Command3.Enabled = True
        Frame3.Enabled = False
        Else
        Command3.Enabled = True
        Frame3.Enabled = False
        End If
    End If
    
    Text13.Text = f
    Frame3.Enabled = False
    
    
    
    
End Sub

Private Sub Form_Load()
    Dim i As Byte
    Dim j As Byte
    Dim k As Byte
    
    Dim data As String
    Dim buffer As String
    Dim char As String
    Dim stab As Integer
    Dim l As Double
    
    l = 0.5
    
    Adodc1.Refresh
    
    Adodc4.Recordset.MoveLast
    u = CInt(Text56.Text)

    
End Sub

Private Sub Timer1_Timer()

    Label16.Caption = Date
    Label17.Caption = Time
    
End Sub
