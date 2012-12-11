VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Comodity 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PT. Angkasa Pura Solusi - CSC Transaction System - Comodity Edit"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Add"
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   480
      Top             =   4680
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Bindings        =   "Comodity.frx":0000
      Height          =   2055
      Left            =   480
      TabIndex        =   6
      Top             =   2400
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3625
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
   Begin VB.TextBox Text2 
      DataField       =   "comdesc"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      DataField       =   "comname"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Description"
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
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Name"
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
      Left            =   480
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Comodity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public statusadmin As Boolean

Private Sub Command1_Click()

    TImbang.statusadmin = True
    TImbang.Show
    Comodity.Enabled = False
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
DataGrid1.Refresh
End If
End Sub

Private Sub Command4_Click()
If Command4.Caption = "Add" Then
Command4.Caption = "Save"
Text1.Enabled = True
Text2.Enabled = True
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
    If Text1.Text = "" Or Text2.Text = "" Then
        Adodc1.Recordset.Cancel
        Adodc1.Refresh
    Else
        Adodc1.Recordset.Update
    End If
Text1.Enabled = False
Text2.Enabled = False
DataGrid1.Refresh
End If
End Sub




