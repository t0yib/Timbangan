VERSION 5.00
Begin VB.Form Comodity 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PT. Angkasa Pura Solusi - CSC Transaction System - Comodity Edit"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   1215
      Left            =   2640
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   4920
      Width           =   1695
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
      Left            =   1080
      TabIndex        =   5
      Top             =   1080
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
      Left            =   1080
      TabIndex        =   4
      Top             =   480
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
