VERSION 5.00
Object = "{24F79378-80CC-4436-9DC7-43217F45C8E3}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Basic Point of Sales System by Adolfo G. Nasol"
   ClientHeight    =   8250
   ClientLeft      =   2370
   ClientTop       =   2460
   ClientWidth     =   14400
   BeginProperty Font 
      Name            =   "Arial Rounded MT Bold"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":628A
   ScaleHeight     =   550
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   960
   Begin VB.Timer tmrTotal 
      Interval        =   1000
      Left            =   480
      Top             =   6120
   End
   Begin VB.CheckBox chkWaterMelon 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      Height          =   270
      Left            =   11160
      TabIndex        =   9
      ToolTipText     =   "Click to Select WaterMelon"
      Top             =   6000
      Width           =   270
   End
   Begin VB.CheckBox chkCherries 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      Height          =   270
      Left            =   7920
      TabIndex        =   8
      ToolTipText     =   "Click to Select Cherries"
      Top             =   6000
      Width           =   270
   End
   Begin VB.CheckBox chkBanana 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      Height          =   270
      Left            =   4560
      TabIndex        =   7
      ToolTipText     =   "Click to Select  Banana"
      Top             =   5880
      Width           =   270
   End
   Begin VB.CheckBox chkGreenApple 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      Height          =   270
      Left            =   1320
      TabIndex        =   6
      ToolTipText     =   "Click to Select Green Apple"
      Top             =   5880
      Width           =   270
   End
   Begin VB.CheckBox chkGreenGrapes 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      Height          =   255
      Left            =   11160
      TabIndex        =   5
      ToolTipText     =   "Click to Select Green Grapes"
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox chkCitron 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      Height          =   255
      Left            =   8160
      TabIndex        =   4
      ToolTipText     =   "Click to Select Citron Fruit"
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox chkAvocado 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      Height          =   270
      Left            =   4560
      TabIndex        =   3
      ToolTipText     =   "Click to Select Avocado"
      Top             =   3600
      Width           =   270
   End
   Begin VB.CheckBox chkStrawberry 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      ForeColor       =   &H0000C0C0&
      Height          =   240
      Left            =   1440
      MaskColor       =   &H80000014&
      TabIndex        =   2
      ToolTipText     =   "Click to Select Strawberry"
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   270
   End
   Begin VB.Timer tmrGreetings 
      Interval        =   1000
      Left            =   1200
      Top             =   240
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgCancelOrder 
      Height          =   345
      Left            =   12840
      Top             =   7080
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      Image           =   "frmMain.frx":3B908
      Effects         =   "frmMain.frx":3BF5C
   End
   Begin VB.Label lblSubTotalHolder 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   9720
      TabIndex        =   12
      Top             =   7800
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Line lnDivider 
      BorderColor     =   &H00808080&
      X1              =   152
      X2              =   152
      Y1              =   416
      Y2              =   464
   End
   Begin VB.Label lblOrderInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   720
      Left            =   2520
      TabIndex        =   11
      Top             =   6360
      Width           =   9780
      WordWrap        =   -1  'True
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgCheckout 
      Height          =   345
      Left            =   11280
      Top             =   7080
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   609
      Image           =   "frmMain.frx":3BF74
      Effects         =   "frmMain.frx":3C644
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgCartFull 
      Height          =   360
      Left            =   1680
      Top             =   7080
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   635
      Image           =   "frmMain.frx":3C65C
      Effects         =   "frmMain.frx":3C91D
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgCartEmpty 
      Height          =   390
      Left            =   1680
      Top             =   7080
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   688
      Image           =   "frmMain.frx":3C935
      Effects         =   "frmMain.frx":3CC42
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Top             =   7200
      Width           =   60
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgWaterMelon 
      Height          =   1860
      Left            =   11040
      ToolTipText     =   "Click to view details"
      Top             =   4560
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   3281
      Image           =   "frmMain.frx":3CC5A
      Effects         =   "frmMain.frx":41808
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgCherries 
      Height          =   1875
      Left            =   7800
      ToolTipText     =   "Click to view details"
      Top             =   4560
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   3307
      Image           =   "frmMain.frx":41820
      Effects         =   "frmMain.frx":45097
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgBanana 
      Height          =   1830
      Left            =   4320
      ToolTipText     =   "Click to view details"
      Top             =   4440
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   3228
      Image           =   "frmMain.frx":450AF
      Effects         =   "frmMain.frx":48143
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgGreenApple 
      Height          =   1875
      Left            =   1080
      ToolTipText     =   "Click to view details"
      Top             =   4440
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   3307
      Image           =   "frmMain.frx":4815B
      Effects         =   "frmMain.frx":4CE0C
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgGreenGrapes 
      Height          =   1605
      Left            =   11160
      ToolTipText     =   "Click to view details"
      Top             =   2280
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   2831
      Image           =   "frmMain.frx":4CE24
      Effects         =   "frmMain.frx":5164D
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgCitron 
      Height          =   1770
      Left            =   8160
      ToolTipText     =   "Click to view details"
      Top             =   2160
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   3122
      Image           =   "frmMain.frx":51665
      Effects         =   "frmMain.frx":54DF6
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgAvocado 
      Height          =   1725
      Left            =   4320
      ToolTipText     =   "Click to view details"
      Top             =   2280
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   3043
      Image           =   "frmMain.frx":54E0E
      Effects         =   "frmMain.frx":58CD3
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgStrawberry 
      Height          =   1830
      Left            =   1200
      ToolTipText     =   "Click to view details"
      Top             =   2160
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   3228
      Image           =   "frmMain.frx":58CEB
      Border          =   49344
      Effects         =   "frmMain.frx":5E01A
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   7800
      Width           =   5655
   End
   Begin VB.Label lblUserName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   11730
      TabIndex        =   0
      Top             =   960
      Width           =   75
   End
   Begin VB.Menu menuMenu 
      Caption         =   "&Menu"
      Index           =   0
      Begin VB.Menu menuSave 
         Caption         =   "Save Sales Report"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu menuDailySalesReport 
         Caption         =   "Daily Sales Report"
         Index           =   1
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu menuAbout 
         Caption         =   "About Basic POS"
         Index           =   3
         Shortcut        =   ^A
      End
      Begin VB.Menu menuLicense 
         Caption         =   "Basic POS License"
         Index           =   15
         Shortcut        =   ^L
      End
      Begin VB.Menu menuWebsite 
         Caption         =   "Authors Website"
         Index           =   20
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuDash3 
         Caption         =   "-"
      End
      Begin VB.Menu menuHelp 
         Caption         =   "Basic POS Help"
         Index           =   4
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu menuClosePOS 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu menuExit 
      Caption         =   "&Exit"
      Index           =   10
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declare showGreetings as Boolean, we will use this
' to create blinking effects in label caption with
' the help of timer
Dim showGreetings As Boolean

' Here is our module level variable, we use module level variable
' so that we can accumulate the value of our sub total
Dim curAvocadoOrdered As Currency, curStrawberryOrdered As Currency
Dim curCitronOrdered As Currency, curGreenGrapesOrdered As Currency
Dim curGreenAppleOrdered As Currency, curBananaOrdered As Currency
Dim curCherriesOrdered As Currency, curWaterMelonOrdered As Currency
Dim curTotalOrdered As Currency

'Let's define some string constant that we will use for tooltip
' We don't want to repeat same string again and again
Const txtTipCancel As String = "Click to cancel order for "
Const txtTipBuy As String = "Click to buy "
Const strTotalPriceInfo As String = "The Total price of your ordered "
Const strTotalPriceText As String = "Total Price for ordered "

' This function will allow us to make a link to a website,
' I use this to open my website http://danreb.com if the user hit the
' Author website in the Menu of the application.
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
   "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
   ByVal lpFile As String, ByVal lpParameters As String, _
   ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Function we use for validating the Quantity input field
Public Function CheckNumeric(IntQuantity, Order)
   
   ' Declare a Boolean variable so we can use True and False
   Dim Valid As Boolean
   
   ' Start our validation process
   While Valid = False
   
   IntQuantity = InputBox("Enter how many kilo(s) of " & Order & vbCrLf _
   & " is your order." & vbCrLf & vbCrLf _
   & "The default MINIMUM order is 1 KILO" & vbCrLf & _
   "Maximum order is 300 Kilos." & vbCrLf & _
   "Please enter only Numeric value.", " How many kilogram(s) of " & Order & "?", 1)
           ' Check for Numeric Entry
           If IsNumeric(IntQuantity) Then
             ' Check if entered value is not 0
              If Val(IntQuantity) <> 0 Then
              ' Check if order is not too much and within a capacity
                If Val(IntQuantity) < 301 Then
                ' Check if Negative number is entered
                  If Not Val(IntQuantity) < 0 Then
                  ' If pass in all test, then success!
                    Valid = True
                    CheckNumeric = Val(IntQuantity)
                  ' or show Message to user
                  Else
                  ' Message for Negative value entered
                    MsgBox "Negative Value not allowed", vbInformation, "Negative Value Detected"
                  End If
                Else
                ' Message if Maximum order is hit
                MsgBox "Out of Range, Maximum order is 300 kilos only" & vbCrLf _
                & "We don't have stock of more than 300 kilos of " & UCase(Order), _
                vbInformation, "Out of Range"
                End If
              Else
              ' Message if Zero is entered
                MsgBox " You cannot order 0 Kilos of " & Order, vbInformation, "Minimum order is 1 Kilo"
              End If
              
            Else
            
               If IntQuantity <> "" Then
               ' Message to show if entered value is not Numeric
                  MsgBox " Enter Numeric Values Only", vbInformation, "Not Valid Value"
               Else
                 Valid = False
                 CheckNumeric = Val(IntQuantity = 0)
                 ' We exit the function after getting the value of 0
                 ' this is activated if the user click on Cancel Button
                 Exit Function
               End If
            End If
    Wend

End Function

Private Sub chkBanana_Click()
   
   ' Declare our Variable
   Dim IntQuantity As String
   ' Set the price of our Banana
   Const curBananaPrice As Currency = 40
   
   ' Check if Banana is selected
   If chkBanana.Value = vbChecked Then
   ' change the tooltip text
   chkBanana.ToolTipText = txtTipCancel & "Banana"
   ' Call our function to validate the Quantity entered
   IntQuantity = CheckNumeric(IntQuantity, "Banana")
   
   ' Check if the returned value of funtion is 0
   If IntQuantity = 0 Then
       ' unchecked and cancel the order if zero
       chkBanana.Value = vbUnchecked
       ' then exit sub, nevermine computing anything from zero
       Exit Sub
   End If
   
   ' Here's the process, we compute how many kilo is ordered
   curBananaOrdered = curBananaPrice * IntQuantity
   ' Inform the users of the Total Price for this product
   MsgBox strTotalPriceInfo & "Banana is PHP " & _
   FormatNumber(curBananaOrdered, 2, vbTrue, vbTrue, vbTrue), _
   vbInformation, strTotalPriceText & "Banana"
   
   End If
   
   ' If the user unchecked our products
   If chkBanana.Value = vbUnchecked Then
    ' reset the price of current Banana Ordered
    curBananaOrdered = 0
    ' change also the tooltip text
    chkBanana.ToolTipText = txtTipBuy & "Banana"
    ' and show message to user that the order is canceled
    MsgBox "Your order for Banana has been canceled" _
    & vbCrLf, vbInformation, "Order Canceled"
   End If
   
End Sub

Private Sub chkCherries_Click()

 ' Declare all our variables
   Dim IntQuantity As String

   Const curCherriesPrice As Currency = 140
   
   ' Check if Cherries is selected
   If chkCherries.Value = vbChecked Then
   ' change the tooltip text
   chkCherries.ToolTipText = txtTipCancel & "Cherries"
   ' Call our function to validate the Quantity entered
   IntQuantity = CheckNumeric(IntQuantity, "Cherries")
   
   ' Check if the returned value of funtion is 0
   If IntQuantity = 0 Then
       ' unchecked and cancel the order if zero
       chkCherries.Value = vbUnchecked
       ' then exit sub, nevermine computing anything from zero
       Exit Sub
   End If
   
   ' Here's the process, we compute how many kilo is ordered
   curCherriesOrdered = curCherriesPrice * IntQuantity
   ' Inform the users of the Total Price for this product
   MsgBox strTotalPriceInfo & "Cherries is PHP " & _
   FormatNumber(curCherriesOrdered, 2, vbTrue, vbTrue, vbTrue), _
   vbInformation, strTotalPriceText & "Cherries"
   
   End If
   
   ' If the user unchecked our products
   If chkCherries.Value = vbUnchecked Then
    ' reset the price of current Cherries Ordered
    curCherriesOrdered = 0
    ' change also the tooltip text
    chkCherries.ToolTipText = txtTipBuy & "Cherries"
    ' and show message to user that the order is canceled
    MsgBox "Your order for Cherries has been canceled" _
    & vbCrLf, vbInformation, "Order Canceled"
   End If
   
End Sub

Private Sub chkCitron_Click()

 ' Declare all our variables
   Dim IntQuantity As String

   Const curCitronPrice As Currency = 70
   
    ' Check if Citron is selected
   If chkCitron.Value = vbChecked Then
   ' change the tooltip text
   chkCitron.ToolTipText = txtTipCancel & "Citron"
   ' Call our function to validate the Quantity entered
   IntQuantity = CheckNumeric(IntQuantity, "Citron")
   
   ' Check if the returned value of funtion is 0
   If IntQuantity = 0 Then
       ' unchecked and cancel the order if zero
       chkCitron.Value = vbUnchecked
       ' then exit sub, nevermine computing anything from zero
       Exit Sub
   End If
   
   ' Here's the process, we compute how many kilo is ordered
   curCitronOrdered = curCitronPrice * IntQuantity
   ' Inform the users of the Total Price for this product
   MsgBox strTotalPriceInfo & "Citron is PHP " & _
   FormatNumber(curCitronOrdered, 2, vbTrue, vbTrue, vbTrue), _
   vbInformation, strTotalPriceText & "Citron"
   
   End If
   
   ' If the user unchecked our products
   If chkCitron.Value = vbUnchecked Then
    ' reset the price of current Citron Ordered
    curCitronOrdered = 0
    ' change also the tooltip text
    chkCitron.ToolTipText = txtTipBuy & "Citron"
    ' and show message to user that the order is canceled
    MsgBox "Your order for Citron has been canceled" _
    & vbCrLf, vbInformation, "Order Canceled"
   End If
   
End Sub

Private Sub chkGreenApple_Click()
   
   ' Declare all our variables
   Dim IntQuantity As String

   Const curGreenApplePrice As Currency = 80
   
   ' Check if Green Apple is selected
   If chkGreenApple.Value = vbChecked Then
   ' change the tooltip text
   chkGreenApple.ToolTipText = txtTipCancel & "Green Apple"
   ' Call our function to validate the Quantity entered
   IntQuantity = CheckNumeric(IntQuantity, " Green Apple")
   
   ' Check if the returned value of funtion is 0
   If IntQuantity = 0 Then
       ' unchecked and cancel the order if zero
       chkGreenApple.Value = vbUnchecked
       ' then exit sub, nevermine computing anything from zero
       Exit Sub
   End If
   
   ' Here's the process, we compute how many kilo is ordered
   curGreenAppleOrdered = curGreenApplePrice * IntQuantity
   ' Inform the users of the Total Price for this product
   MsgBox strTotalPriceInfo & "GreenApple is PHP " & _
   FormatNumber(curGreenAppleOrdered, 2, vbTrue, vbTrue, vbTrue), _
   vbInformation, strTotalPriceText & "Green Apple"
   
   End If
   
   ' If the user unchecked our products
   If chkGreenApple.Value = vbUnchecked Then
    ' reset the price of current Green Apple Ordered
    curGreenAppleOrdered = 0
    ' change also the tooltip text
    chkGreenApple.ToolTipText = txtTipBuy & "Green Apple"
    ' and show message to user that the order is canceled
    MsgBox "Your order for Green Apple has been canceled" _
    & vbCrLf, vbInformation, "Order Canceled"
   End If
   
End Sub

Private Sub chkGreenGrapes_Click()
   
   ' Declare all our variables
   Dim IntQuantity As String

   Const curGreenGrapesPrice As Currency = 100
   
   ' Check if Green Grapes is selected
   If chkGreenGrapes.Value = vbChecked Then
   ' change the tooltip text
   chkGreenGrapes.ToolTipText = txtTipCancel & "Green Grapes"
   ' Call our function to validate the Quantity entered
   IntQuantity = CheckNumeric(IntQuantity, " Green Grapes")
   
   ' Check if the returned value of funtion is 0
   If IntQuantity = 0 Then
       ' unchecked and cancel the order if zero
       chkGreenGrapes.Value = vbUnchecked
       ' then exit sub, nevermine computing anything from zero
       Exit Sub
   End If
   
   ' Here's the process, we compute how many kilo is ordered
   curGreenGrapesOrdered = curGreenGrapesPrice * IntQuantity
   ' Inform the users of the Total Price for this product
   MsgBox strTotalPriceInfo & " Green Grapes is PHP " & _
   FormatNumber(curGreenGrapesOrdered, 2, vbTrue, vbTrue, vbTrue), _
   vbInformation, strTotalPriceText & "Green Grapes"
   
   End If
   
   ' If the user unchecked our products
   If chkGreenGrapes.Value = vbUnchecked Then
    ' reset the price of current Green Grapes Ordered
    curGreenGrapesOrdered = 0
    ' change also the tooltip text
    chkGreenGrapes.ToolTipText = txtTipBuy & "Green Grapes"
    ' and show message to user that the order is canceled
    MsgBox "Your order for Green Grapes has been canceled" _
    & vbCrLf, vbInformation, "Order Canceled"
   End If
   
End Sub

Private Sub chkStrawberry_Click()
   
   ' Declare all our variables
   Dim IntQuantity As String

   Const curStrawberryPrice As Currency = 80
   
   ' Check if Strawberry is selected
   If chkStrawberry.Value = vbChecked Then
   ' change the tooltip text
   chkStrawberry.ToolTipText = txtTipCancel & "Strawberry"
   ' Call our function to validate the Quantity entered
   IntQuantity = CheckNumeric(IntQuantity, "Strawberry")
   
   ' Check if the returned value of funtion is 0
   If IntQuantity = 0 Then
       ' unchecked and cancel the order if zero
       chkStrawberry.Value = vbUnchecked
       ' then exit sub, nevermine computing anything from zero
       Exit Sub
   End If
   
   ' Here's the process, we compute how many kilo is ordered
   curStrawberryOrdered = curStrawberryPrice * IntQuantity
   ' Inform the users of the Total Price for this product
   MsgBox strTotalPriceInfo & "Strawberry is PHP " & _
   FormatNumber(curStrawberryOrdered, 2, vbTrue, vbTrue, vbTrue), _
   vbInformation, strTotalPriceText & "Strawberry"
   
   End If
   
   ' If the user unchecked our products
   If chkStrawberry.Value = vbUnchecked Then
    ' reset the price of current Strawberry Ordered
    curStrawberryOrdered = 0
    ' change also the tooltip text
    chkStrawberry.ToolTipText = txtTipBuy & "Strawberry"
    ' and show message to user that the order is canceled
    MsgBox "Your order for Strawberry has been canceled" _
    & vbCrLf, vbInformation, "Order Canceled"
   End If
   
End Sub

Private Sub chkWaterMelon_Click()
   
   ' Declare all our variables
   Dim IntQuantity As String

   Const curWaterMelonPrice As Currency = 90
   
   ' Check if Water Melon is selected
   If chkWaterMelon.Value = vbChecked Then
   ' change the tooltip text
   chkWaterMelon.ToolTipText = txtTipCancel & "Water Melon"
   ' Call our function to validate the Quantity entered
   IntQuantity = CheckNumeric(IntQuantity, "Water Melon")
   
   ' Check if the returned value of funtion is 0
   If IntQuantity = 0 Then
       ' unchecked and cancel the order if zero
       chkWaterMelon.Value = vbUnchecked
       ' then exit sub, nevermine computing anything from zero
       Exit Sub
   End If
   
   ' Here's the process, we compute how many kilo is ordered
   curWaterMelonOrdered = curWaterMelonPrice * IntQuantity
   ' Inform the users of the Total Price for this product
   MsgBox strTotalPriceInfo & "WaterMelon is PHP " & _
   FormatNumber(curWaterMelonOrdered, 2, vbTrue, vbTrue, vbTrue), _
   vbInformation, strTotalPriceText & "Water Melon"
   
   End If
   
   ' If the user unchecked our products
   If chkWaterMelon.Value = vbUnchecked Then
    ' reset the price of current Water Melon Ordered
    curWaterMelonOrdered = 0
    ' change also the tooltip text
    chkWaterMelon.ToolTipText = txtTipBuy & "Water Melon"
    ' and show message to user that the order is canceled
    MsgBox "Your order for Water Melon has been canceled" _
    & vbCrLf, vbInformation, "Order Canceled"
   End If
   
End Sub

Private Sub Form_Load()
  
  ' Center the Main form in the screen
   Dim frm As Form
    For Each frm In Forms
        frm.Left = (Screen.Width / 2) - (frm.Width / 2)
        frm.Top = (Screen.Height / 2) - (frm.Height / 2)
    Next
  
  ' We create another variable that we will use to greet the logged in user
  Dim GreetUser As String
  GreetUser = "Greetings " & StrConv(frmLogin.txtUserName.Text, vbProperCase) _
  & ", have a Good day!"
  lblUserName.Caption = GreetUser
  ' Set showGreetings to true to show it
  showGreetings = True
  ' Get the current date
  lblDate.Caption = "Today's Date: " & Format(Now, "MMM dd YYYY")
  
End Sub

Private Sub chkAvocado_Click()

   ' Declare all our variables
   Dim IntQuantity As String
   
   Const curAvocadoPrice As Currency = 50
   
   ' Check if Avocado is selected
   If chkAvocado.Value = vbChecked Then
   ' change the tooltip text
   chkAvocado.ToolTipText = txtTipCancel & "Avocado"
   ' Call our function to validate the Quantity entered
   IntQuantity = CheckNumeric(IntQuantity, "Avocado")
   
   ' Check if the returned value of funtion is 0
   If IntQuantity = 0 Then
       ' unchecked and cancel the order if zero
       chkAvocado.Value = vbUnchecked
       ' then exit sub, nevermine computing anything from zero
       Exit Sub
   End If
   
   ' Here's the process, we compute how many kilo is ordered
   curAvocadoOrdered = curAvocadoPrice * IntQuantity
   ' Inform the users of the Total Price for this product
   MsgBox strTotalPriceInfo & "Avocado is PHP " & _
   FormatNumber(curAvocadoOrdered, 2, vbTrue, vbTrue, vbTrue), _
   vbInformation, strTotalPriceText & "Avocado"
   
   End If
   
   ' If the user unchecked our products
   If chkAvocado.Value = vbUnchecked Then
    ' reset the price of current Avocado Ordered
    curAvocadoOrdered = 0
    ' change also the tooltip text
    chkAvocado.ToolTipText = txtTipBuy & "Avocado"
    ' and show message to user that the order is canceled
    MsgBox "Your order for Avocado has been canceled" _
    & vbCrLf, vbInformation, "Order Canceled"
   End If
   
End Sub

Private Sub imgAvocado_MouseEnter()

   ' Change the opacity of the fruit on mouse hover
   imgAvocado.TransparencyPct = 30
   
End Sub

Private Sub imgAvocado_MouseExit()

  ' Change back to nomal opacity when the mouse exit
   imgAvocado.TransparencyPct = 0

End Sub

Private Sub imgBanana_MouseEnter()

   ' Change the design and opacity of the fruit on mouse hover
   imgBanana.TransparencyPct = 30
   
End Sub

Private Sub imgBanana_MouseExit()

   ' Change back to nomal opacity when the mouse exit
   imgBanana.TransparencyPct = 0
    
End Sub

Private Sub imgCancelOrder_Click()
    
    ' We unload the Main form
    Unload frmMain
    Unload frmCheckout
    ' And inform the user about resetting and canceling of order
    MsgBox "All your order is being canceled" _
    & " and the Cart is being reset.", vbInformation, "Reset Order Information"
    
    ' Then we will reset to zero all our module level variable
    curTotalOrdered = 0
    curAvocadoOrdered = 0
    curStrawberryOrdered = 0
    curCitronOrdered = 0
    curGreenGrapesOrdered = 0
    curGreenAppleOrdered = 0
    curBananaOrdered = 0
    curCherriesOrdered = 0
    curWaterMelonOrdered = 0
    
    ' And after reset to zero our variable, we will again show the main form
    frmMain.Show
    
End Sub

Private Sub imgCancelOrder_MouseEnter()
    
    imgCancelOrder.GrayScale = lvicSepia
    
End Sub

Private Sub imgCancelOrder_MouseExit()
   
   imgCancelOrder.GrayScale = lvicNoGrayScale
   
End Sub

Private Sub imgCheckout_Click()

' We will use this later to compute our Total Price
  frmCheckout.Show
  
End Sub

Private Sub imgCheckout_MouseEnter()

  ' Change the color of the button on mouse hover
   imgCheckout.GrayScale = lvicSepia
   
End Sub

Private Sub imgCheckout_MouseExit()
   
   ' Set the color to normal
   imgCheckout.GrayScale = lvicNoGrayScale
   
End Sub

Private Sub imgCherries_MouseEnter()

   ' Change the opacity of the fruit on mouse hover
   imgCherries.TransparencyPct = 30

End Sub

Private Sub imgCherries_MouseExit()

   ' Change back to nomal opacity when the mouse exit
   imgCherries.TransparencyPct = 0

End Sub

Private Sub imgCitron_MouseEnter()

   ' Change the opacity of the fruit on mouse hover
  imgCitron.TransparencyPct = 30

End Sub

Private Sub imgCitron_MouseExit()
   
   ' Change back to nomal opacity when the mouse exit
   imgCitron.TransparencyPct = 0
    
End Sub

Private Sub imgGreenApple_MouseEnter()

  ' Change the opacity of the fruit on mouse hover
   imgGreenApple.TransparencyPct = 30
   
End Sub

Private Sub imgGreenApple_MouseExit()

   ' Change back to nomal opacity when the mouse exit
  imgGreenApple.TransparencyPct = 0

End Sub

Private Sub imgGreenGrapes_MouseEnter()

  ' Change the opacity of the fruit on mouse hover
  imgGreenGrapes.TransparencyPct = 30
   
End Sub

Private Sub imgGreenGrapes_MouseExit()

   ' Change back to nomal opacity when the mouse exit
   imgGreenGrapes.TransparencyPct = 0
    
End Sub

Private Sub imgStrawberry_MouseEnter()

   ' Change the opacity of the fruit on mouse hover
   imgStrawberry.TransparencyPct = 30

End Sub

Private Sub imgStrawberry_MouseExit()
   
   ' Change back to nomal opacity when the mouse exit
   imgStrawberry.TransparencyPct = 0
    
End Sub

Private Sub imgWaterMelon_MouseEnter()

   ' Change the opacity of the fruit on mouse hover
   imgWaterMelon.TransparencyPct = 30
   
End Sub

Private Sub imgWaterMelon_MouseExit()

   ' Change back to nomal opacity when the mouse exit
  imgWaterMelon.TransparencyPct = 0
    
End Sub

Private Sub menuAbout_Click(Index As Integer)

   'Show the about form
   frmAbout.Show
   
End Sub

Private Sub menuClosePOS_Click()
    
    ' End the program
    ' I want to make sure all form is unloaded first.
    Unload frmMain
    Unload frmAbout
    Unload frmCheckout
    Unload frmLogin
    Unload frmSplash
    End
    
End Sub

Private Sub menuDailySalesReport_Click(Index As Integer)
  
  ' Declare variable that will hold our Message Confirmation to users.
  Dim strConfirm As String
  
 ' We will open the folder for Sales Reports so users
 ' can view the existing Daily Sales Reports, but we will ask for
 ' confirmation first.
  strConfirm = MsgBox("Hit OK to open the Sales Report folder" & vbCrLf _
   & "Press the Cancel Button if you want me to cancel this action.", _
  vbOKCancel, "Sales Reports Viewing Confirmation")
  
  If strConfirm = vbOK Then
    Shell "Explorer.exe " & frmCheckout.lblSalesReportsHolder.Caption, vbNormalFocus
  End If

End Sub

Private Sub menuExit_Click(Index As Integer)

  'Unload the form
  Unload Me
  ' End the program
  End
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
 
 ' I want to make sure the about form is unloaded
 Unload frmAbout
 
 ' and all values of our Module level variable is reset to zero
 curAvocadoOrdered = 0
 curStrawberryOrdered = 0
 curCitronOrdered = 0
 curGreenGrapesOrdered = 0
 curGreenAppleOrdered = 0
 curBananaOrdered = 0
 curCherriesOrdered = 0
 curWaterMelonOrdered = 0
 curTotalOrdered = 0

End Sub

Private Sub menuHelp_Click(Index As Integer)
   
   MsgBox "TODO, I need to create the Help File.", vbInformation, "Help File coming soon...."
   
End Sub

Private Sub menuLicense_Click(Index As Integer)
   
   ' Release my Basic Point of Sales System as Freeware.
   MsgBox "I released this software as FREEWARE, you are Free to study the code" & vbCrLf _
   & "that is used in this program, thanks to my brother Google because, " & vbCrLf _
   & "He guide's me throughout the whole process on creating this Application" & vbCrLf _
   & vbCrLf & "I fully commented the code so that everyone can easily identify what" & vbCrLf _
   & "I am trying to do, if you find some bugs or you feel we can still improve it" & vbCrLf _
   & "Please send me an email: adolfo@danreb.com , you can also visit " & vbCrLf _
   & "my website http://www.danreb.com" & vbCrLf & vbCrLf _
   & "Distributing my Application in your website is permitted but" & vbCrLf _
   & "you will provide a link to http://www.danreb.com" & vbCrLf & vbCrLf _
   & "Basic Point of Sale system ® Author : Adolfo G. Nasol." & vbCrLf _
   & "Copyright © 2011: Adolfo G. Nasol", vbInformation, "License Information"
   
End Sub

Private Sub menuSave_Click(Index As Integer)
    
   ' Check if we have an existing order before we save anything to our
   ' sales reports file.
   If curTotalOrdered <> 0 Then
     MsgBox "Sales Summary, make sure the ordered product(s) is paid" & vbCrLf _
     & "Before saving the data to Daily Sales Reports.", vbInformation, "Sales Summary"
     frmCheckout.Show
   Else
     MsgBox "We don't have sales data to save, your Cart is empty." _
     , vbInformation, "No Sales Data to Save!"
   End If

End Sub

Private Sub menuWebsite_Click(Index As Integer)
    
    ' Open my website
    ShellExecute Me.hwnd, "Open", "http://danreb.com", 0, 0, 3
    
End Sub

Private Sub tmrGreetings_Timer()
  
  ' Here we use this code to toggle the visibility
  ' of Greetings Message for logged in user
  If showGreetings = True Then
    lblUserName.Visible = True
    showGreetings = False
 Else
    lblUserName.Visible = False
    showGreetings = True
 End If
 
End Sub

Private Sub tmrTotal_Timer()
     
     ' We will use this timer to automaticcaly update
     ' our Sub Total price in Every 2 seconds
     
     curTotalOrdered = curAvocadoOrdered + curStrawberryOrdered + _
     curCitronOrdered + curGreenGrapesOrdered + curGreenAppleOrdered _
     + curBananaOrdered + curCherriesOrdered + curWaterMelonOrdered
     
     lblSubTotalHolder.Caption = curTotalOrdered
     
     ' If total Sales is 0, we will hide the Total Price
     If curTotalOrdered = 0 Then
     ' and will inform the users that Cart is empty
        lblTotal.Caption = "Cart is empty, no products in the Cart."
        imgCheckout.Visible = False
        imgCartEmpty.Visible = True
        imgCartFull.Visible = False
        lblOrderInfo.Visible = False
        imgCancelOrder.Visible = False
     Else
     ' If there is an order we will Show the Sub Total Price
     ' computation for all the selected and ordered products
        lblTotal.Caption = "Sub Total : Php " & FormatNumber(curTotalOrdered, 2, True, True, True)
        imgCheckout.Visible = True
        imgCartEmpty.Visible = False
        imgCartFull.Visible = True
        lblOrderInfo.Visible = True
        imgCancelOrder.Visible = True
     End If
     
     ' We will use this to show the current order of the user
     ' All order will be concatinated to string variable strInfo
     Dim strInfo As String
     strInfo = "Your Order: "
     Dim curAvocadoOrderedQuantity As Currency
     Dim curStrawberryOrderedQuantity As Currency
     Dim curCitronOrderedQuantity As Currency
     Dim curGreenGrapesOrderedQuantity As Currency
     Dim curGreenAppleOrderedQuantity As Currency
     Dim curBananaOrderedQuantity As Currency
     Dim curCherriesOrderedQuantity As Currency
     Dim curWaterMelonOrderedQuantity As Currency
     
     ' Check if there is an order for Avocado
     If curAvocadoOrdered <> 0 Then
        curAvocadoOrderedQuantity = curAvocadoOrdered / 50
        strInfo = strInfo & curAvocadoOrderedQuantity & " Kilo(s) of Avocado, "
     End If
     
     ' Check if there is an order for Strawberry
     If curStrawberryOrdered <> 0 Then
        curStrawberryOrderedQuantity = curStrawberryOrdered / 80
        strInfo = strInfo & curStrawberryOrderedQuantity & " Kilo(s) of Strawberry, "
     End If
     
     ' Check if there is an order for Citron
     If curCitronOrdered <> 0 Then
        curCitronOrderedQuantity = curCitronOrdered / 70
        strInfo = strInfo & curCitronOrderedQuantity & " Kilo(s) of Citron, "
     End If
     
     ' Check if there is an order for Green Grapes
     If curGreenGrapesOrdered <> 0 Then
         curGreenGrapesOrderedQuantity = curGreenGrapesOrdered / 100
         strInfo = strInfo & curGreenGrapesOrderedQuantity & " Kilo(s) of Green Grapes, "
     End If
     
     ' Check if there is an order for Green Apple
     If curGreenAppleOrdered <> 0 Then
         curGreenAppleOrderedQuantity = curGreenAppleOrdered / 80
         strInfo = strInfo & curGreenAppleOrderedQuantity & " Kilo(s) of Green Apple, "
     End If
          
      ' Check if there is an order for Banana
     If curBananaOrdered <> 0 Then
         curBananaOrderedQuantity = curBananaOrdered / 40
         strInfo = strInfo & curBananaOrderedQuantity & " Kilo(s) of Banana, "
     End If
     
     ' Check if there is an order for Cherries
     If curCherriesOrdered <> 0 Then
         curCherriesOrderedQuantity = curCherriesOrdered / 140
         strInfo = strInfo & curCherriesOrderedQuantity & " Kilo(s) of Cherries, "
     End If
     
     ' Check if there is an order for Water Melon
     If curWaterMelonOrdered <> 0 Then
         curWaterMelonOrderedQuantity = curWaterMelonOrdered / 90
         strInfo = strInfo & curWaterMelonOrderedQuantity & " Kilo(s) of WaterMelon, "
     End If
     
     ' We will then get all the result and show it as caption for label
     lblOrderInfo.Caption = strInfo
     
End Sub
