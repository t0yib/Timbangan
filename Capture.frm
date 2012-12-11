VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   4680
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6360
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   0
      Top             =   3600
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Byte
Dim j As Byte
Dim k As Byte
Dim l As Double
Dim data As String
Dim buffer As String
Dim char As String

l = 0.5
End Sub

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
                        Text2.Text = j
                        Text3.Text = k
                            For i = j To k
                            char = Mid(buffer, i, 1)
                            If (char = "1") Or (char = "2") Or (char = "3") Or (char = "4") Or (char = "5") Or (char = "6") Or (char = "7") Or (char = "8") Or (char = "9") Or (char = "0") Or (char = "-") Then
                            data = data + char
                        End If
                        i = i + l
                        l = l * 2
                        Next i
                        Text1.Text = data
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

