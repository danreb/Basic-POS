VERSION 5.00
Object = "{24F79378-80CC-4436-9DC7-43217F45C8E3}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Authentication"
   ClientHeight    =   4500
   ClientLeft      =   3570
   ClientTop       =   3300
   ClientWidth     =   7515
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":628A
   ScaleHeight     =   2658.749
   ScaleMode       =   0  'User
   ScaleWidth      =   7056.178
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEnter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Default         =   -1  'True
      Height          =   135
      Left            =   7560
      TabIndex        =   2
      Top             =   4440
      Width           =   15
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      ToolTipText     =   "Enter your User name"
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      DragIcon        =   "frmLogin.frx":DB8E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "Enter your Password"
      Top             =   2040
      Width           =   3165
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgError 
      Height          =   435
      Left            =   6240
      Top             =   2040
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   767
      Image           =   "frmLogin.frx":13E18
      Effects         =   "frmLogin.frx":1471B
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgCheck 
      Height          =   360
      Left            =   6240
      Top             =   2040
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   635
      Image           =   "frmLogin.frx":14733
      Effects         =   "frmLogin.frx":14D63
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnCancelShadow 
      Height          =   375
      Left            =   4200
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Image           =   "frmLogin.frx":14D7B
      Settings        =   153600
      Effects         =   "frmLogin.frx":15341
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnLoginShadow 
      Height          =   375
      Left            =   2880
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Image           =   "frmLogin.frx":15359
      Settings        =   153600
      Effects         =   "frmLogin.frx":15867
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnSendCancel 
      Height          =   375
      Left            =   4200
      ToolTipText     =   "Click to Cancel"
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Image           =   "frmLogin.frx":1587F
      Effects         =   "frmLogin.frx":15E45
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl btnSendLogin 
      Height          =   375
      Left            =   2880
      ToolTipText     =   "Click to Login"
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Image           =   "frmLogin.frx":15E5D
      Effects         =   "frmLogin.frx":1636B
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  ' First we Set the LoginSucceded variable as Boolean so that
  ' we can use true or false in login action
Public LoginSucceeded As Boolean

' Then we try to Get the UserName of current user using API
' and we will automatically asign this name to the textbox for username
Private Declare Function GetUserName Lib "advapi32.dll" _
Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

' Function to get the current logged on user in windows
Public Function CurrentUser() As String

    ' We create a string buffer of 255 chars in length
    Dim strBuff As String * 255
    ' We declare X and use it to hold length of username
    Dim x As Long
    CurrentUser = ""
    x = GetUserName(strBuff, Len(strBuff) - 1)
    
    ' We check then if x is greater that 0 in length
    If x > 0 Then
        'We look for Null Character also, usually included
        x = InStr(strBuff, vbNullChar)
        'Trim off buffered spaces too
        If x > 0 Then
            CurrentUser = Left$(strBuff, x - 1)
        Else
            CurrentUser = Left$(strBuff, x)
        End If
    End If
    
End Function

Private Sub btnSendCancel_Click()

    'set the global var to false to denote a failed login
    LoginSucceeded = False
    ' and unload the login form
    Unload Me
    
End Sub

Private Sub btnSendCancel_MouseEnter()

   ' Change the color of the button on mouse hover
   btnSendCancel.GrayScale = lvicSepia
   ' Also adjust the shadow to reflect the changes
   With btnCancelShadow
    .TransparencyPct = 60
    .GrayScale = lvicSepia
   End With
   
End Sub

Private Sub btnSendCancel_MouseExit()
  
  ' On mouse Exit, make the button in its original state
  btnSendCancel.GrayScale = lvicNoGrayScale
  ' Also adjust the shadow to reflect the changes
   With btnCancelShadow
    .TransparencyPct = 88
    .GrayScale = lvicNoGrayScale
   End With
  
End Sub

Private Sub btnSendLogin_MouseEnter()

   ' Change the color of the button on mouse hover
   btnSendLogin.GrayScale = lvicSepia
   ' Also adjust the shadow to reflect the changes
   With btnLoginShadow
    .TransparencyPct = 60
    .GrayScale = lvicSepia
   End With
  
End Sub

Private Sub btnSendLogin_MouseExit()
  
  ' On mouse Exit, make the button in its original state
  btnSendLogin.GrayScale = lvicNoGrayScale
  ' Also adjust the shadow to reflect the changes
   With btnLoginShadow
    .TransparencyPct = 88
    .GrayScale = lvicNoGrayScale
   End With
  
End Sub

Private Sub btnSendLogin_Click()

    'check for correct password
    If txtPassword.Text = "ict-tr3a1" Then
        LoginSucceeded = True
        ' show the Main form
        frmMain.Show
        ' and unload the login form
        Unload Me
    Else
        ' Tell the user about incorrect password
         MsgBox "Invalid Password, Please try again!" & vbCrLf _
        & " Make sure you did not activate CapsLock", vbInformation, "Invalid Password Information"
        ' Clear the password field
        txtPassword.Text = ""
        ' Then set focus to password field
        txtPassword.SetFocus
    End If
    
End Sub

Private Sub cmdEnter_Click()
    ' This command button is hidden
    ' I only use this so that users can still Press Enter
    ' in the keyboard to submit the login form
    ' please note that this code is exactly the same
    ' with the code for image button above
    'check for correct password
    If txtPassword.Text = "ict-tr3a1" Then
        LoginSucceeded = True
        ' show the Main form
        frmMain.Show
        ' and unload the login form
        Unload Me
    Else
        ' Tell the user about incorrect password
        MsgBox "Invalid Password, Please try again!" & vbCrLf _
        & " Make sure you did not activate CapsLock", vbInformation, "Invalid Password Information"
        ' Clear the password field
        txtPassword.Text = ""
        ' Then set focus to password field
        txtPassword.SetFocus
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' This snippet will create an unload effects in form unload events
  Dim i As Long
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
  For i = Me.Left To (Screen.Width / 2) Step 2
     Me.Height = Me.Height - 250
     Me.Width = Me.Width - 250
     Me.Left = Me.Left + 25
     Me.Refresh
     DoEvents
  Next
   Unload Me
 
End Sub

Private Sub Form_Load()

    ' We now assign the name of logged in user in windows
    ' to our Textbox in login form
    txtUserName.Text = CurrentUser
    
End Sub

Private Sub txtPassword_Change()

     ' Use the change events in password fields to give hints to user
     ' we just swap the visibility of the two image here
  If txtPassword.Text = "ict-tr3a1" Then
    ' If password is correct show the check icon and hide the
    ' cross icon
      imgError.Visible = False
      imgCheck.Visible = True
    Else
    ' If the password is not correct we show the cross icon
    ' and we hide the check icon
      imgError.Visible = True
      imgCheck.Visible = False
  End If
  
End Sub
