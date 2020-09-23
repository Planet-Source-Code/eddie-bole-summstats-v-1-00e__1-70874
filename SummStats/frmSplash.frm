VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2865
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6135
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timTimer 
      Interval        =   155
      Left            =   95
      Top             =   1800
   End
   Begin VB.Label LblBy 
      BackStyle       =   0  'Transparent
      Caption         =   " Original code created by Anton Venema.       Additional features added by Eddie Bole         PDF Producer written by Paras Chopra"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   1200
      TabIndex        =   3
      Top             =   1920
      Width           =   4095
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A simple program to check your results"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5895
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Summary Statistics Finder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   -120
      TabIndex        =   1
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Always ""Freeware"" GPL, Version 1.00e Build 25"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   -120
      TabIndex        =   0
      Top             =   1440
      Width           =   6375
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer
Private Sub Form_Load()

    'Set the default value for i
    i = (-1)

End Sub

Private Sub timTimer_Timer()
Dim j As Integer

    'Increase i
    i = i + 1

    Select Case i

        'Load the form
        Case 10
            Load frmStatistics

        'Display the form and close the splash screen
        Case 12
            frmStatistics.Show
            Unload Me
    End Select

End Sub
