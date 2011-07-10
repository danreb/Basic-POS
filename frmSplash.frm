VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Welcome to Basic POS system"
   ClientHeight    =   5400
   ClientLeft      =   3600
   ClientTop       =   3945
   ClientWidth     =   8250
   ControlBox      =   0   'False
   DrawMode        =   16  'Merge Pen
   Enabled         =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSplash 
      Interval        =   8000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Author: Mr. Adolfo G. Nasol
' ICT - TR3A1

' First set this option on so that we need to declare a variable
Option Explicit

' We use this code to embed an executable file inside our form
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

' Use the API function to find and load the exe window
  Private Declare Function FindWindow Lib "user32" _
  Alias "FindWindowA" (ByVal lpClassName As String, _
  ByVal lpWindowName As String) As Long

  Private Declare Function SetParent Lib "user32" _
  (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

  Private Declare Function SetWindowPos Lib "user32" _
  (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
  ByVal cy As Long, ByVal wFlags As Long) As Long

  Private Declare Function GetClientRect Lib "user32" _
  (ByVal hwnd As Long, lpRect As RECT) As Long

' We use this code to make our form transparent
  Private Declare Function GetWindowLong Lib "user32" _
  Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

  Private Declare Function SetWindowLong Lib "user32" _
  Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
               
  Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hwnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long
                
  Private Const GWL_STYLE = (-16)
  Private Const GWL_EXSTYLE = (-20)
  Private Const WS_EX_LAYERED = &H80000
  Private Const LWA_COLORKEY = &H1
  Private Const LWA_ALPHA = &H2
  
  ' We use this code to close the opened executable files.
  Private Const WM_CLOSE As Long = &H10
  Private Declare Function SendMessage Lib "user32" _
  Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
  ByVal wParam As Long, lParam As Any) As Long
  
Private Sub Form_Load()
  
  ' Declare our variable that will hold our object
  Dim ProjectIntro As Long, frmParent As RECT
  
  ' Call our executable file
  Call Shell(App.Path & "\intro.exe", vbNormalFocus)
  
  ' Assign the exe file to our variable
  ProjectIntro = FindWindow(vbNullString, "Welcome to Basic POS system")
  
  ' Then we use the getclient function to
  ' make our form as parent window of the exe file
  Call GetClientRect(Me.hwnd, frmParent)
  Call SetParent(ProjectIntro, Me.hwnd)
  
  ' then we set the position of the exe file inside the parent window
   Call SetWindowPos(ProjectIntro, 0&, frmParent.Left, frmParent.Top, _
  frmParent.Right, frmParent.Bottom, 1)
  
   ' Then make our form transparent so it can't be seen
    Me.BackColor = vbCyan
    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, _
    GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, vbCyan, 0&, LWA_COLORKEY
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim ProjectIntro As Long
    ProjectIntro = FindWindow(vbNullString, "Welcome to Basic POS system")
    'Then Unload the Flash Projector Intro
    Call SendMessage(ProjectIntro, WM_CLOSE, 0, 0)
    
End Sub

Private Sub tmrSplash_Timer()

    ' We use the timer to automatically show the Login form
    frmLogin.Show
    ' Then we unload this form
    Unload frmSplash
    
End Sub


