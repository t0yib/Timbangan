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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   7680
      Top             =   1080
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
      Height          =   555
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000B&
      Caption         =   "FIX"
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   6720
      TabIndex        =   2
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RUN"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   3480
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
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   2280
      Width           =   4215
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   8160
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      SThreshold      =   1
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "NOT FIX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   1695
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
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "KG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   2400
      Width           =   495
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
      Width           =   1215
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
Dim m As Integer
Dim data As String
Dim buffer As String
Dim char As String


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

Private Sub Command3_Click()
    m = 0
    Cls
    Timer1.Interval = 3000
    Timer1.Enabled = True
    
    'Do Until (Timer1.Interval = 5000)
        'm = n
    'Loop
        
    'If m Is n Then
        'MsgBox (m)
        'Label4.Caption = "FIXED"
        'Label4.BackColor = &HFF00&
        'Label4.ForeColor = &H80000012
    'Else
        'MsgBox ("none")
        'Label4.Caption = "NOT FIX"
        'Label4.BackColor = &HFF&
        'Label4.ForeColor = &H8000000B
    'End If
        
End Sub

Private Sub Timer1_Timer()

    m = m + 1
    
End Sub
