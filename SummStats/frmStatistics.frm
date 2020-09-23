VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmStatistics 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Summary Statistics Finder And Worksheet Maker"
   ClientHeight    =   9660
   ClientLeft      =   30
   ClientTop       =   630
   ClientWidth     =   12735
   FillColor       =   &H00FFC0C0&
   Icon            =   "frmStatistics.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   644
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   849
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox sorted 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   56
      Top             =   2880
      Width           =   5775
   End
   Begin VB.CommandButton cmd_createworksheet 
      BackColor       =   &H00FF8080&
      Caption         =   "Save Worksheet, F9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Clear all results"
      Top             =   8520
      Width           =   1935
   End
   Begin VB.CommandButton clr_results 
      BackColor       =   &H00FF8080&
      Caption         =   "Clear All Results, F12"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Clear all results"
      Top             =   8520
      Width           =   1815
   End
   Begin VB.CommandButton quit_but 
      BackColor       =   &H00FFC0C0&
      DownPicture     =   "frmStatistics.frx":09CA
      Height          =   930
      Left            =   5040
      Picture         =   "frmStatistics.frx":1BA6
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Quit the program"
      Top             =   360
      Width           =   945
   End
   Begin VB.CommandButton clear_but 
      BackColor       =   &H00FFC0C0&
      DownPicture     =   "frmStatistics.frx":2C9D
      Height          =   930
      Left            =   3840
      Picture         =   "frmStatistics.frx":3F8F
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Clear the results"
      Top             =   360
      Width           =   945
   End
   Begin VB.CommandButton analyse_but 
      BackColor       =   &H00FFC0C0&
      DownPicture     =   "frmStatistics.frx":5173
      Height          =   930
      Left            =   2640
      Picture         =   "frmStatistics.frx":638C
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Analyse the data values"
      Top             =   360
      Width           =   945
   End
   Begin VB.CommandButton save_but 
      BackColor       =   &H00FFC0C0&
      DownPicture     =   "frmStatistics.frx":7533
      Height          =   930
      Left            =   1440
      Picture         =   "frmStatistics.frx":86B3
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Save the results"
      Top             =   360
      Width           =   945
   End
   Begin VB.CommandButton open 
      BackColor       =   &H00FFC0C0&
      DownPicture     =   "frmStatistics.frx":97DE
      Height          =   930
      Left            =   240
      Picture         =   "frmStatistics.frx":A91E
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Open a data file"
      Top             =   360
      Width           =   945
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3015
      Left            =   6600
      TabIndex        =   47
      ToolTipText     =   "The Summary Statistics"
      Top             =   1800
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmStatistics.frx":BA69
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton SaveAll 
      BackColor       =   &H00FF8080&
      Caption         =   "Save The Results, F8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Save the summary statistics"
      Top             =   8520
      Width           =   1815
   End
   Begin VB.TextBox txtStat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   19
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox txtStat 
      BeginProperty DataFormat 
         Type            =   5
         Format          =   ""
         HaveTrueFalseNull=   1
         TrueValue       =   "True"
         FalseValue      =   "False"
         NullValue       =   ""
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   7
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   18
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox txtStat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   17
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   7920
      Width           =   1575
   End
   Begin VB.TextBox txtStat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   16
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton CloseAll 
      BackColor       =   &H00FF8080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Quit the program"
      Top             =   8520
      Width           =   1935
   End
   Begin VB.CommandButton cmdAnalyse 
      BackColor       =   &H00FF8080&
      Caption         =   "Analyse, F5"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Analyse the results"
      Top             =   8520
      Width           =   1815
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00FF8080&
      Caption         =   "ClearAbove, F7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Clear the above values"
      Top             =   8520
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog cdgList 
      Left            =   6480
      Top             =   8520
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      FileName        =   "values.txt"
      Filter          =   "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
      InitDir         =   "App.Path"
   End
   Begin VB.TextBox txtStat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   8
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   33
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Frame fraList 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Data List (delimited by commas)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   6015
      Begin VB.TextBox txtList 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         OLEDropMode     =   2  'Automatic
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         ToolTipText     =   "Enter your data values & click Analyse ( F5 )"
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.TextBox txtStat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   14
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   5040
      Width           =   1575
   End
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   30
      Top             =   9285
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "  Enter list delimited by commas"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtStat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   15
      Left            =   4440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Top             =   7920
      Width           =   1575
   End
   Begin VB.TextBox txtStat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   13
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txtStat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   12
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtStat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   11
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox txtStat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   10
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtStat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   9
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txtStat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   7
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   7200
      Width           =   1575
   End
   Begin VB.TextBox txtStat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   6
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   7200
      Width           =   1575
   End
   Begin VB.TextBox txtStat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   5
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txtStat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   4
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox txtStat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   3
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox txtStat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   2
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox txtStat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   1
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox txtStat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Graphical Controls"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   120
      TabIndex        =   48
      Top             =   120
      Width           =   6015
   End
   Begin RichTextLib.RichTextBox wshtquest 
      Height          =   3015
      Left            =   6600
      TabIndex        =   60
      ToolTipText     =   "The worksheet questions"
      Top             =   5280
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmStatistics.frx":BAE4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblWorksheetquestions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Worksheet Questions:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   6480
      TabIndex        =   59
      Top             =   5040
      Width           =   2040
   End
   Begin VB.Label lblResults 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Results:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   6480
      TabIndex        =   58
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label lblSorted 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sorted Data Values:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   20
      Left            =   120
      TabIndex        =   57
      Top             =   2640
      Width           =   1515
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      BorderWidth     =   5
      X1              =   0
      X2              =   848
      Y1              =   616
      Y2              =   616
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      BorderWidth     =   10
      X1              =   0
      X2              =   848
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      BorderWidth     =   3
      X1              =   424
      X2              =   424
      Y1              =   104
      Y2              =   600
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Upper Outlier Present"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   19
      Left            =   2160
      TabIndex        =   45
      Top             =   5520
      Width           =   1680
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lower Outlier Present"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   18
      Left            =   120
      TabIndex        =   44
      Top             =   5520
      Width           =   1680
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Upper Outlier Calculation"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   17
      Left            =   2160
      TabIndex        =   41
      Top             =   7680
      Width           =   1965
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lower Outlier Calculation"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   16
      Left            =   120
      TabIndex        =   40
      Top             =   7680
      Width           =   1965
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Error:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   8
      Left            =   4320
      TabIndex        =   34
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Range:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   4320
      TabIndex        =   32
      Top             =   4080
      Width           =   510
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deviations:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   15
      Left            =   4320
      TabIndex        =   25
      Top             =   7680
      Width           =   870
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode(s):"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   4320
      TabIndex        =   24
      Top             =   3360
      Width           =   660
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interquartile Range:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   14
      Left            =   4320
      TabIndex        =   23
      Top             =   4800
      Width           =   1560
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Variance"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   2160
      TabIndex        =   22
      Top             =   6960
      Width           =   690
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Q3:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   12
      Left            =   2160
      TabIndex        =   10
      Top             =   4800
      Width           =   300
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Max X:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   13
      Left            =   2160
      TabIndex        =   11
      Top             =   4080
      Width           =   555
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Median:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   11
      Left            =   2160
      TabIndex        =   9
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Q1:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   10
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   270
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Min X:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   9
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   510
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sample Size (n):"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   6240
      Width           =   1260
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Deviation:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   7
      Left            =   120
      TabIndex        =   5
      Top             =   6960
      Width           =   1545
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sum of X²:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   4320
      TabIndex        =   4
      Top             =   6240
      Width           =   870
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sum of X:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   2160
      TabIndex        =   3
      Top             =   6240
      Width           =   765
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mean:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   465
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileAnalyse 
         Caption         =   "&Analyse"
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFileClear 
         Caption         =   "Clear The &LHS Results"
         Shortcut        =   {F7}
      End
      Begin VB.Menu wipeall 
         Caption         =   "&Clear All Results"
         Shortcut        =   {F12}
      End
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu SaveAllAnswers 
         Caption         =   "&Save The Results"
         Shortcut        =   {F8}
      End
      Begin VB.Menu savequestions 
         Caption         =   "Save &Worksheet Questions"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save &Data List"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open List"
         Shortcut        =   ^O
      End
      Begin VB.Menu separator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "C&lose"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu Tools 
      Caption         =   "&Tools"
      Begin VB.Menu convert2pdf 
         Caption         =   "Convert To &PDF"
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu about 
      Caption         =   "About"
      Index           =   0
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "&Quit"
   End
   Begin VB.Menu by 
      Caption         =   "                           Various sections programmed and designed  by Anton Venema, Paras Chopra and Eddie Bole (2000 - 2008)"
      NegotiatePosition=   1  'Left
   End
End
Attribute VB_Name = "frmStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dataArray() As Long
Private objHeapSort As New clsHeapSort


Private Sub about_Click(Index As Integer)
    frmSplash.Show
    frmSplash.Refresh
End Sub

Private Sub analyse_but_Click()
mnuFileAnalyse_Click
End Sub



Private Sub by_Click()
    frmSplash.Show
    frmSplash.Refresh
End Sub

Private Sub clear_but_Click()
mnuFileClear_Click
End Sub

Private Sub CloseAll_Click()
    'End program
    Unload Me
    End
End Sub

Private Sub clr_results_Click()
If RichTextBox1.TextRTF <> "" Then
        RichTextBox1.TextRTF = ""
        wshtquest.TextRTF = ""
  End If
  
  mnuFileClear_Click
  txtList.SetFocus
  
End Sub

Private Sub cmd_createworksheet_Click()
On Error GoTo ErrorHandler
Dim FileNumber As Long

    'Display Save Dialog Box
    cdgList.ShowSave

    'Write test box contents to file
    FileNumber = FreeFile
    Open cdgList.filename For Output As #FileNumber
        Write #FileNumber, wshtquest.Text
    Close #FileNumber

    'Display status of program
    stbStatus.SimpleText = "List data written to file: " + cdgList.filename

    'Exit before error handler
    Exit Sub

ErrorHandler:

    'Cancel was pressed
    If Err.Number = 32755 Then

        'Display status of program
        stbStatus.SimpleText = "Data not written to file"
        Exit Sub
    End If

End Sub

Private Sub cmdAnalyse_Click()

    'Call the clicking of the corresponding menu item
    mnuFileAnalyse_Click

End Sub

Private Sub cmdClear_Click()

    'Call the clicking of the corresponding menu item
    mnuFileClear_Click

End Sub

Private Sub convert2pdf_Click()
formConvertToPDF.Show
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    'F5 = Analyse button
    If KeyCode = vbKeyF5 Then
        If cmdAnalyse.Enabled = True Then
            cmdAnalyse_Click
        End If

    'F7 = Clear button
    ElseIf KeyCode = vbKeyF7 Then
        cmdClear_Click
    End If

 

End Sub

Private Sub Form_Load()

    'Set form as always on top
    'SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, SWP_SHOWWINDOW

    'Increase window count
    WindowCount = WindowCount + 1

End Sub

Private Sub Form_Unload(Cancel As Integer)

    End
    
End Sub

Private Sub mnuFileAnalyse_Click()
Dim i As Long

    'Show status bar text
    stbStatus.SimpleText = "Working..."

    'Set basic array dimensions
    ReDim MainArray(0)

    'Set basic variables for ExtractToken
    CurrentEntryIndex = 1
    InputString = txtList.Text

    'Get the first token and check to make
    'sure it is a number
    ExtractToken
    If OutputString = "Number" Then
        MainArray(0) = OutputValue
    Else
        ErrorHandler
        Exit Sub
    End If

    'Loop until end of string, checking for
    'correct placement of commas and numbers
    Do While CurrentEntryIndex < Len(InputString)

        'Check for comma
        If Mid(InputString, CurrentEntryIndex, 1) <> "," Then
            ErrorHandler
        End If

        'Check for number
        CurrentEntryIndex = CurrentEntryIndex + 1
        ExtractToken
        If OutputString = "Number" Then
            ReDim Preserve MainArray(UBound(MainArray) + 1)
            MainArray(UBound(MainArray)) = OutputValue
        Else
            ErrorHandler
            Exit Sub
        End If
    Loop

    stbStatus.SimpleText = ""

    'Call each statistic function
    txtStat(0).Text = Round(Mean, 5)
    txtStat(1).Text = Mode
    txtStat(2).Text = SumX
    txtStat(3).Text = SumX2
    txtStat(4).Text = SS
    txtStat(5).Text = Range
    txtStat(6).Text = Round(Variance, 5)
    txtStat(7).Text = Round(SD, 5)
    txtStat(8).Text = "±" + CStr(SE)
    txtStat(9).Text = minX
    txtStat(10).Text = Q1
     
    txtStat(11).Text = Median
    txtStat(12).Text = Q3
    txtStat(13).Text = maxX
    txtStat(14).Text = IR
    txtStat(15).Text = Deviations

    txtStat(16).Text = LowOut
    txtStat(17).Text = UppOut

    txtStat(18).Text = L_Out
    txtStat(19).Text = U_Out
    sort_Click
    
'RichTextBox1.SelText = "Numbers:" + Chr(13) + Chr(10) + txtList.Text + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "Mean = " + txtStat(0).Text + ",  Median = " + txtStat(11).Text + ",  Mode = " + txtStat(1).Text + Chr(13) + Chr(10) + "MinX = " + txtStat(9).Text + ",  MaxX = " + txtStat(13).Text + ",  Range = " + txtStat(5).Text + Chr(13) + Chr(10) + "Q1 = " + txtStat(10).Text + ",  Q2 = " + txtStat(12).Text + ",  IQR = " + txtStat(14).Text + Chr(13) + Chr(10) + "Lower Outlier = " + txtStat(18).Text + ",  Upper Outlier = " + txtStat(19).Text + Chr(13) + Chr(10) + Chr(13) + Chr(10) & vbCrLf
RichTextBox1.SelText = "Numbers: " & txtList.Text & Chr(13) & Chr(10) & "In  Order: " & sorted.Text & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Mean = " & txtStat(0).Text & ",  Median = " & txtStat(11).Text & ",  Mode = " & txtStat(1).Text & Chr(13) & Chr(10) & "MinX = " & txtStat(9).Text & ",  MaxX = " & txtStat(13).Text & ",  Range = " & txtStat(5).Text & Chr(13) & Chr(10) & "Q1 = " & txtStat(10).Text & ",  Q2 = " & txtStat(12).Text & ",  IQR = " & txtStat(14).Text & Chr(13) & Chr(10) & "Lower Outlier = " & txtStat(18).Text & ",  Upper Outlier = " & txtStat(19).Text & Chr(13) & Chr(10) & Chr(13) & Chr(10) & vbCrLf & Chr(10) & vbCrLf

wshtquest.SelText = "Numbers: " & txtList.Text & Chr(10) & vbCrLf & "In  Order: " & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Mean = " & vbTab & ",  Median = " & vbTab & vbTab & ",  Mode = " & vbTab & Chr(10) & vbCrLf & "MinX = " & vbTab & vbTab & ",  MaxX = " & vbTab & vbTab & ",  Range = " & Chr(10) & vbCrLf & "Q1 = " & vbTab & vbTab & ",  Q2 = " & vbTab & vbTab & vbTab & ",  IQR = " & Chr(10) & vbCrLf & "Lower Outlier = " & vbTab & vbTab & vbTab & ",  Upper Outlier = " & vbTab & Chr(13) & Chr(10) & Chr(13) & Chr(10) & vbCrLf & Chr(10) & vbCrLf




    'Show status bar text
    stbStatus.SimpleText = "Done"

End Sub


Private Sub mnuFileClear_Click()
Dim i As Integer

    'Clear list box first, and then the answers
    If txtList.Text <> "" Then
        txtList.Text = ""
        sorted.Text = ""
        txtStat(0).Text = ""
        txtStat(1).Text = ""
        txtStat(2).Text = ""
        txtStat(3).Text = ""
        txtStat(4).Text = ""
        txtStat(5).Text = ""
        txtStat(6).Text = ""
        txtStat(7).Text = ""
        txtStat(8).Text = ""
        txtStat(9).Text = ""
        txtStat(10).Text = ""
        txtStat(11).Text = ""
        txtStat(12).Text = ""
        txtStat(13).Text = ""
        txtStat(14).Text = ""
        txtStat(15).Text = ""
        txtStat(16).Text = ""
        txtStat(17).Text = ""
        txtStat(18).Text = ""
        txtStat(19).Text = ""
  End If

End Sub

Private Sub mnuFileClose_Click()
    
    'Close Statistics form
    Unload Me

End Sub

Private Sub mnuFileOpen_Click()
On Error GoTo ErrorHandler
Dim FileNumber As Long
Dim ListString As String

    'Display Open dialog box
    cdgList.ShowOpen

    'Get data from file into text box
    ListString = ""
    FileNumber = FreeFile
    Open cdgList.filename For Input As #FileNumber
        Do While Not EOF(FileNumber)
            Input #FileNumber, InputString
            ListString = ListString + InputString
        Loop
    Close #FileNumber
    txtList.Text = ListString
    txtList.SetFocus
    txtList.SelStart = Len(txtList.Text)

    'Display status of program
    stbStatus.SimpleText = "List data received from file: " + cdgList.filename

    'Exit before error handler
    Exit Sub

ErrorHandler:

    'File not found
    If Err.Number = 53 Then

        'Display error message
        MsgBox "Error: File not found", vbInformation, "File not found"
        stbStatus.SimpleText = "File not found"
        Exit Sub
    End If

End Sub

Private Sub mnuFileSave_Click()
On Error GoTo ErrorHandler
Dim FileNumber As Long

    'Display Save Dialog Box
    cdgList.ShowSave

    'Write test box contents to file
    FileNumber = FreeFile
    Open cdgList.filename For Output As #FileNumber
        Write #FileNumber, txtList.Text
    Close #FileNumber

    'Display status of program
    stbStatus.SimpleText = "List data written to file: " + cdgList.filename

    'Exit before error handler
    Exit Sub

ErrorHandler:

    'Cancel was pressed
    If Err.Number = 32755 Then

        'Display status of program
        stbStatus.SimpleText = "List data not written to file"
        Exit Sub
    End If

End Sub

Private Sub mnuQuit_Click()
    'End program
    Unload Me
    End
End Sub

Private Sub open_Click()
mnuFileOpen_Click
End Sub















Private Sub quit_but_Click()
 'End program
    Unload Me
    End
End Sub

Private Sub save_but_Click()
    SaveAllAnswers_Click
End Sub

Private Sub SaveAllAnswers_Click()
On Error GoTo ErrorHandler
Dim FileNumber As Long

    'Display Save Dialog Box
    cdgList.ShowSave

    'Write test box contents to file
    FileNumber = FreeFile
    Open cdgList.filename For Output As #FileNumber
        Write #FileNumber, RichTextBox1.Text
    Close #FileNumber

    'Display status of program
    stbStatus.SimpleText = "List data written to file: " + cdgList.filename

    'Exit before error handler
    Exit Sub

ErrorHandler:

    'Cancel was pressed
    If Err.Number = 32755 Then

        'Display status of program
        stbStatus.SimpleText = "List data not written to file"
        Exit Sub
    End If
End Sub

Private Sub SaveAll_Click()

SaveAllAnswers_Click
    
End Sub





Private Sub savequestions_Click()
cmd_createworksheet_Click
End Sub

Private Sub txtList_Change()
Dim i As Long
Dim Count As Long

    'Disable Analyse button if no entered text
    If txtList.Text = "" Then
        cmdAnalyse.Enabled = False
        mnuFileAnalyse.Enabled = False
        mnuFileSave.Enabled = False

    'Enable Analyse button if text box contains something
    Else
        cmdAnalyse.Enabled = True
        mnuFileAnalyse.Enabled = True
        mnuFileSave.Enabled = True
    End If

End Sub

Private Sub ErrorHandler()

    'Display error message
    stbStatus.SimpleText = "Syntax Error"

End Sub


Private Sub sort_Click()
    Dim i As Long
    Dim Temp As Variant
    
    Temp = Split(txtList.Text, ",")
    ReDim dataArray(1 To 1)
    For i = 0 To UBound(Temp)
        If IsNumeric(Temp(i)) Then
            dataArray(UBound(dataArray)) = CLng(Temp(i))
            ReDim Preserve dataArray(1 To UBound(dataArray) + 1)
        End If
    Next i
    
    ReDim Preserve dataArray(1 To UBound(dataArray) - 1)
    Call objHeapSort.heapSort(dataArray)
    sorted.Text = ""
    For i = 1 To UBound(dataArray)
        sorted.Text = sorted.Text & dataArray(i) & ","
    Next i
End Sub

Private Sub wipeall_Click()
    clr_results_Click
End Sub
