VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar"
      Height          =   495
      Left            =   6720
      TabIndex        =   2
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Jalankan"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox display 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2040
      Width           =   4215
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4080
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "TIMBANGAN DISPLAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   5
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "KG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "BERAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Byte
Dim j As Byte
Dim k As Byte
Dim data As String
Dim buffer As String
Dim char As String
Dim m As Double

Private Sub Command1_Click()

    MSComm1.PortOpen = True
        Do
        DoEvents
            data = MSComm1.Input
            buffer = buffer + data
            i = Len(buffer)
            For j = 1 To i
                char = Mid(buffer, j, 1)
                    If char = "S" Then
                        For k = j To i
                        char = Mid(buffer, k, 1)
                            If char = Chr(13) Then
                            data = ""
                                For i = j To k
                                char = Mid(buffer, i, 1)
                                    If (char = "1") Or (char = "2") Or (char = "3") Or (char = "4") Or (char = "5") Or (char = "6") Or (char = "7") Or (char = "8") Or (char = "9") Or (char = "0") Or (char = "-") Then
                                    data = data + char
                                    End If
                                i = i + 1
                                Next i
                            display.Text = data
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
    
    Result = MsgBox("Close Program?", vbYesNo)
    If Result = vbYes Then
    End
    End If
    
End Sub


