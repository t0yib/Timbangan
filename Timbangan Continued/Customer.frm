VERSION 5.00
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
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   375
      Left            =   7560
      TabIndex        =   11
      Top             =   6840
      Width           =   1575
   End
   Begin VB.TextBox Text10 
      Height          =   975
      Left            =   7320
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   7320
      TabIndex        =   9
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   855
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   855
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   9360
      TabIndex        =   0
      Top             =   6840
      Width           =   1695
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
