VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Timbang 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PT. Angkasa Pura Solusi - CSC Transaction System"
   ClientHeight    =   7860
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   15300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   15300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "RUN"
      Height          =   375
      Left            =   11160
      TabIndex        =   30
      Top             =   6960
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000003&
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
      Left            =   1680
      TabIndex        =   24
      Top             =   1560
      Width           =   4095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "+ CSC No."
      BeginProperty Font 
         Name            =   "Bell Gothic Std Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   23
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox display 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3960
      Top             =   6360
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8040
      TabIndex        =   16
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8040
      TabIndex        =   12
      Top             =   2880
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save / Print"
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   6960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   8040
      TabIndex        =   9
      Text            =   " . . . "
      Top             =   1560
      Width           =   6975
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6000
      TabIndex        =   7
      Text            =   "Pilih Flight"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "[ + ] Add / Edit Customer"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   2520
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   600
      TabIndex        =   3
      Text            =   "Pilih Customer"
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   13200
      TabIndex        =   0
      Top             =   6960
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
      TabIndex        =   5
      Top             =   2160
      Width           =   5535
      Begin VB.Frame Frame2 
         BackColor       =   &H80000005&
         Caption         =   "Consigner Data View"
         Height          =   2535
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   5055
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5400
         Y1              =   960
         Y2              =   960
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Commodity Data View"
      Height          =   4455
      Left            =   6000
      TabIndex        =   13
      Top             =   2040
      Width           =   9135
      Begin VB.CommandButton Command4 
         Caption         =   "Edit Commodity"
         Height          =   375
         Left            =   6840
         TabIndex        =   31
         Top             =   360
         Width           =   1695
      End
      Begin VB.Frame Frame4 
         Height          =   2295
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Width           =   8895
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Add Commodity"
         Height          =   375
         Left            =   4920
         TabIndex        =   28
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   19
         Top             =   1560
         Width           =   2655
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   2040
         TabIndex        =   18
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "AWB / SMU Code"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "SPCL Code"
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantity"
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Commodity"
         Height          =   255
         Left            =   840
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
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
      Left            =   1560
      TabIndex        =   27
      Top             =   1080
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
      Left            =   360
      TabIndex        =   26
      Top             =   1080
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
      Left            =   6600
      TabIndex        =   25
      Top             =   120
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   6480
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
      TabIndex        =   22
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label LabelTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      TabIndex        =   21
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "Remarks"
      Height          =   255
      Left            =   8040
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Flight No / Info"
      Height          =   255
      Left            =   6000
      TabIndex        =   8
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
      Left            =   8280
      TabIndex        =   1
      Top             =   120
      Width           =   6855
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   7200
      Top             =   120
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   7200
      Shape           =   2  'Oval
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape4 
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   8160
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   7095
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   3495
   End
End
Attribute VB_Name = "TImbang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConnMySql As New ADODB.Connection
Dim RsMySql As New ADODB.Recordset
Dim CmdMySql As New ADODB.Command
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

Private Sub Command6_Click()

    Text6.Enabled = True
    Text6.BackColor = &H80000005
    Combo1.Enabled = True
    Combo2.Enabled = True
    

End Sub

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
                        For k = 1 To i
                        char = Mid(buffer, k, 1)
                            If char = Chr(13) Then
                            data = ""
                                For i = j To k
                                char = Mid(buffer, i, 1)
                                    If (char = "1") Or (char = "2") Or (char = "3") Or (char = "4") Or (char = "5") Or (char = "6") Or (char = "7") Or (char = "8") Or (char = "9") Or (char = "0") Or (char = "-") Then
                                    data = data + char
                                    End If
                                i = i + l
                                l = l * 2
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
    Login.statusadmin = True
    Login.Show
    Unload Me
    End If
    
End Sub

Private Sub Timer1_Timer()

    LabelTime.Caption = Time
    LabelDate.Caption = Date
    
End Sub
