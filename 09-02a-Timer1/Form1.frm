VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   2445
   ClientTop       =   1905
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   6375
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop Timer"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   3060
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Timer"
      Height          =   495
      Left            =   420
      TabIndex        =   0
      Top             =   3060
      Width           =   1215
   End
   Begin VB.Timer tmrTest 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5580
      Top             =   3060
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintCount As Integer

Private Sub cmdStart_Click()

    mintCount = 0
    Cls
    tmrTest.Enabled = True

End Sub


Private Sub cmdStop_Click()

    tmrTest.Enabled = False

End Sub


Private Sub tmrTest_Timer()

    mintCount = mintCount + 1

    Print "Timer fired again. Count = " & mintCount
    
End Sub
