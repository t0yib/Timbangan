VERSION 5.00
Begin VB.Form Login 
   Caption         =   "PT. Angkasa Pura Solusi - CSC Transaction System - Login"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8895
   LinkTopic       =   "Form2"
   ScaleHeight     =   4680
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   5
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3840
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Operator Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public statusadmin As Boolean

Private Sub Command1_Click()

Call Koneksi
    If Text1.Text = "" Then
        MsgBox "NAMA USER MASIH KOSONG !", vbCritical + vbOKOnly, "Error"
        Text1.SetFocus
        
    ElseIf Text2.Text = "" Then
        MsgBox "PASSWORD MASIH KOSONG !", vbCritical + vbOKOnly, "Error"
        Text2.SetFocus
    
    Else
        SQL = ""
        SQL = "SELECT * FROM user " _
            & "WHERE username='" & Text1.Text & "' " _
            & " AND userpass='" & Text2.Text & "'"
            
            Set rsPeriksa = conn.Execute(SQL)
                   
        If Not rsPeriksa.BOF Then
            TImbang.statusadmin = True
            TImbang.Show
            Unload Me
            
        Else
                MsgBox "ANDA BUKAN USER YANG BERHAK!", vbCritical + vbOKOnly, "Error"
        End If
    End If

End Sub

Private Sub Command2_Click()

    Result = MsgBox("Close Program?", vbYesNo)
    
    If Result = vbYes Then
    End
    End If
    
End Sub

Private Sub Form_Load()

Text1.Text = ""
Text2.Text = ""
End Sub
