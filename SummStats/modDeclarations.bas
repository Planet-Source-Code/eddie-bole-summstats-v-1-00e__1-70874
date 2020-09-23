Attribute VB_Name = "modDeclarations"
Option Explicit

'General
Public Char As String
Public CurrentEntryIndex As Integer
Public DecIndex As Integer
Public ErrorMessage As String
Public Help As Boolean
Public InError As Boolean
Public InputString As String
Public OutputString As String
Public OutputValue As Double
Public PrevAnswer As Double
Public PrevEntry As String
Public SetVariable As Boolean
Public Value As Double
Public ValueString As String
Public WindowCount As Integer

'Variable Storing
Public MainArray() As Double
Public ValueArray() As Double
Public VariableArray() As String

'Constants
  Public Const Pi = 3.14159265358979

'Always on top window display
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1
