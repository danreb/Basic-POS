VERSION 5.00
Begin VB.Form frmCheckout 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Checkout, Calculate and Save Sales Reports"
   ClientHeight    =   6705
   ClientLeft      =   4080
   ClientTop       =   2610
   ClientWidth     =   7740
   ForeColor       =   &H00000000&
   Icon            =   "frmCheckout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNextOrder 
      Caption         =   "&Next Order"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   12
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   11
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   10
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdCalculateChanges 
      BackColor       =   &H80000009&
      Caption         =   "Accept &Payment"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   7
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox txtCashPayment 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      ToolTipText     =   "Please Enter the Amount"
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Label lblSalesReportsHolder 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   7560
      TabIndex        =   13
      Top             =   120
      Width           =   135
   End
   Begin VB.Label lblCustomerChange 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   3885
      TabIndex        =   9
      Top             =   5760
      Width           =   90
   End
   Begin VB.Label lblCashPaid 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Width           =   1170
   End
   Begin VB.Line lnChangeDivider 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   3960
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label lblCashPaid 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PHP"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   4920
      Width           =   570
   End
   Begin VB.Label lblCashPaid 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Payment Amount :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   2955
   End
   Begin VB.Line lnLineTotal 
      BorderColor     =   &H00C0C0C0&
      X1              =   120
      X2              =   3960
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label lblTotalSum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   75
   End
   Begin VB.Label lblVatTaxes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   75
   End
   Begin VB.Label lblSubTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   75
   End
   Begin VB.Line lnHorDivider 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   3960
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line lnVertDivider 
      BorderColor     =   &H00C0C0C0&
      X1              =   3960
      X2              =   3960
      Y1              =   0
      Y2              =   2760
   End
   Begin VB.Label lblReports 
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
      ForeColor       =   &H00808080&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmCheckout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 ' Declarations of Module level variable, we need this for the creation of sales reports file
 ' inside the Application Data folder
Public Enum eSpecialFolders
  SpecialFolder_AppData = &H1A        'for the current Windows user, on any computer on the network [Windows 98 or later]
  SpecialFolder_CommonAppData = &H23  'for all Windows users on this computer [Windows 2000 or later]
  SpecialFolder_LocalAppData = &H1C   'for the current Windows user, on this computer only [Windows 2000 or later]
  SpecialFolder_Documents = &H5       'the Documents folder for the current Windows user
End Enum

  'This function will check the Application Data folder
Public Function SpecialFolder(pFolder As eSpecialFolders) As String

Dim objShell  As Object
Dim objFolder As Object

  Set objShell = CreateObject("Shell.Application")
  Set objFolder = objShell.namespace(CLng(pFolder))
  If (Not objFolder Is Nothing) Then SpecialFolder = objFolder.Self.Path
  Set objFolder = Nothing
  Set objShell = Nothing
  If SpecialFolder = "" Then Err.Raise 513, "SpecialFolder", "The folder path could not be detected"

End Function

Private Sub cmdCalculateChanges_Click()
    
    ' Declare all our variable
    Dim curReportsSubTotal As Currency, curReportsVatTaxes As Currency
    Dim curReportsSumTotal As Currency, curCustomerChange As Currency
    
    ' Make our variable holds and process certain value
    ' curReportsSubTotal will hold the value of Sub Total
    ' From the Main Form.
    curReportsSubTotal = Val(frmMain.lblSubTotalHolder.Caption)
    
    ' We will then compute Vat Taxes 12% of total order
    curReportsVatTaxes = curReportsSubTotal * 0.12
    ' Then we will add Tax Plus the total amount of the ordered produtcs
    curReportsSumTotal = curReportsSubTotal + curReportsVatTaxes
    
    ' We will now go on and validate the value in the Cash Payment form
    ' Let's check if there is value entered in the text box
       If txtCashPayment.Text <> "" Then
       ' If there is a value and the users use a comma, we need to remove
       ' the comma so that we can calculate much better.
          txtCashPayment.Text = Replace(txtCashPayment.Text, ",", "")
          ' Then we need to check if the entered value is too much
          If Val(txtCashPayment.Text) > 500000 Then
          ' If it is too much we will show a message informing the users
          ' of the maximum amount we can accept.
            MsgBox "Out of Capacity, We can only make change for" & vbCrLf _
            & "Money that is less than Php 500,000", vbInformation, "Out of Range or Capacity."
          ' we then reset the textbox field and set it on focus
            txtCashPayment.Text = "0.00"
            txtCashPayment.SetFocus
          ' We need to check again and make sure the freak users
          ' will not pay an amount less than the total amount needs to be paid.
          ElseIf Val(txtCashPayment.Text) < curReportsSumTotal Then
          ' Show a message informing the users for the error.
            MsgBox "Credit is good but we need Cash." & vbCrLf _
            & "It seem's you entered a negative value OR" & vbCrLf _
            & "An amount which is less than the required cash" & vbCrLf _
            & "for this payment.", vbInformation, "Cash Payment Only!"
           ' Do the resetting of textbox again
            txtCashPayment.Text = "0.00"
            txtCashPayment.SetFocus
          Else
           ' If all the above validation checking pass,
           ' we will compute the change
           curCustomerChange = Val(txtCashPayment.Text) - curReportsSumTotal
           ' then we will show it to users in the label caption
           lblCustomerChange.Caption = "Php " & _
           FormatNumber(curCustomerChange, 2, vbTrue, vbTrue, vbTrue)
          End If
       Else
       ' If the Cash Payment is blank, we will inform the users
       ' to enter some amount that we can compute
        MsgBox "Please Enter the Amount in the Text Box" & vbCrLf _
        & " So I can compute the changes.", vbInformation, "No Cash Amount paid!"
        ' and set the value to 0.00 to signify that it must be currency or
        ' numeric value
        txtCashPayment.Text = "0.00"
        ' and we focus the user to Cash Payment text field again
        txtCashPayment.SetFocus
       End If
            
End Sub

Private Sub cmdClear_Click()
     
     ' If the user click on Clear button, we will replace the value of
     ' the Cash Payment text box with 0.00 signifying currency
     txtCashPayment.Text = "0.00"
     ' and we will clear the value of Customer Change
     lblCustomerChange.Caption = ""
     ' and set the focus to Cash Payment textbox again
     txtCashPayment.SetFocus
     
End Sub

Private Sub cmdNextOrder_Click()

    Dim ClearOrder As String
    
    ' If the users click on the Next Order button
    ' We will first ask if he/she want to save the sales report
    ' so that there will be no sales information that will lost accidentally.
    '(Maliban na lang kung gusto talaga mangupit ng cashier.)
    ' We will unload the Main form if OK is click ,
    ' and this will reset all the data in the Main Form
    ClearOrder = MsgBox("Please make sure you save the sales" & vbCrLf _
    & "before you go on to your Next Order." & vbCrLf _
    & "Click OK to continue Or Cancel to SAVE" _
    & " the Sales Report first." _
    , vbOKCancel, "Make sure to save the Sales Report.")
    
    If ClearOrder = vbOK Then
      Unload frmMain
      ' Immediately we will also unload this check out form
      Unload frmCheckout
      ' Then reload the Main form fresh and clear
      frmMain.Show
    End If
    
End Sub

Private Sub cmdSave_Click()

    'Declare all our variables that we will manipulate later
    Dim strSales As String, intFile As Integer
    Dim StrDate As String, StrReportDate As String
    Dim strReports As String, strTaxCollected As Currency
    Dim strTotalSales As Currency, strGrossSales As Currency
    Dim strPath As String

    ' Initializition, assign default value to some of our variables declare above
    strSales = ""
    StrDate = Format(Now, "mmmm-d-yyyy")
    StrReportDate = Format(Now, "mmmm d yyyy at hh:mm AM/PM")
    strTotalSales = frmMain.lblSubTotalHolder.Caption
    strTaxCollected = frmMain.lblSubTotalHolder.Caption * 0.12
    strGrossSales = strTotalSales + strTaxCollected
    strGrossSales = FormatNumber(strGrossSales, 2, vbTrue, vbTrue, vbTrue)
    strPath = SpecialFolder(SpecialFolder_CommonAppData) & "\BasicPOS\SalesReports\"
     
    ' strReports is the variable that we use to hold data that we write to text file
    ' if the user hit the save button
    ' We need to format our Text Report, you can see how we format it this code.
    strReports = Replace(frmMain.lblOrderInfo.Caption, "Your Order:", _
    vbCrLf & vbCrLf & "Products Sales for: " & StrReportDate _
    & vbCrLf & vbCrLf)
    strSales = "**********************************************"
    strSales = strSales & Replace(strReports, ",", vbCrLf)
    strSales = strSales & "______________________________________________" _
    & vbCrLf & vbCrLf & "  Total Sales   : Php " & _
    FormatNumber(strTotalSales, 2, vbTrue, vbTrue, vbTrue) & vbCrLf
    strSales = strSales & "  Tax Collected : Php " & _
    FormatNumber(strTaxCollected, 2, vbTrue, vbTrue, vbTrue) & vbCrLf
    strSales = strSales & "  Gross Sales   : Php " & _
    FormatNumber(strGrossSales, 2, vbTrue, vbTrue, vbTrue) _
    & vbCrLf & vbCrLf & vbCrLf _
    & "**********************************************"
    
    ' We will create and write the reports in the text file
    ' First let us get a Free or unused file by assigning a Free File to our variable.
    intFile = FreeFile
    ' Then we need to open the text file inside the reports folder,
    ' if it exist we will write and append the sales data in the text file.
    ' If the text file doesn't exist, our program will automatically create that text file.
    ' The only requirement was there is an existing SalesReports folder, if there is no folder
    ' our program will return an error ( Todo - I need to researh , what is the work around on this problem. )

    Open strPath & StrDate & "-Sales.txt" _
    For Append As #intFile
    Write #intFile, strSales
    Close #intFile
    
    ' Inform the users about the saved file
    MsgBox "Sales data was saved to text file located in: " & _
    vbCrLf & strPath & StrDate & "-Sales.txt" & _
    vbCrLf & vbCrLf, vbInformation, "Sales Reports Information"
    
    ' Then we will open that text file so that users
    ' can print or edit the sales reports.
    
    Dim myRetVal
    Dim strOpenSalesReport As String
    
    strOpenSalesReport = "notepad " & strPath & StrDate & "-Sales.txt"
    myRetVal = Shell(strOpenSalesReport, vbMaximizedFocus)
    

End Sub

Private Sub Form_Load()
  
  ' Declare all our variable first immediately after the Check out form load
  Dim strReports As String, strReportsSubTotal As String
  Dim curReportsSubTotal As Currency, curTotalSumCharges As Currency
  Dim strPath As String
  strPath = SpecialFolder(SpecialFolder_CommonAppData)
  ' We need to use this constant for our VAT Taxes of 12%
  Const curReportsVatTaxes = 0.12
  
  ' Let's grab the value of Sub Total amount from the Main form
  ' so that we can use it later in the computation
  curReportsSubTotal = Val(frmMain.lblSubTotalHolder.Caption)
  
  ' We also need to grab the value of Order Information from the Main form
  ' so we can show a summary of order report to users in this check out form
  ' We also need to remove the comma and replace it with line break
  ' to make the summary for each ordered fruit to show each in every line
  ' we use the Replace() and vbCrLf to make this happen
  strReports = Replace(frmMain.lblOrderInfo.Caption, "Your Order:", "Summary of Ordered Products: " _
  & vbCrLf & vbCrLf)
  lblReports.Caption = Replace(strReports, ",", vbCrLf)
  
  ' Let also grab the value from the Main form for formatted SubTotal amount
  ' and show to user in our summary information section
  lblSubTotal.Caption = frmMain.lblTotal.Caption
  
  ' Lets compute the tax charge and show it to users
  lblVatTaxes.Caption = "Vat Taxes: Php " & FormatNumber(curReportsVatTaxes * curReportsSubTotal, 2, True, True, True)
  ' Then add the Tax and Order Total so the customer can see
  ' the total amount he/she need to pay.
  lblTotalSum.Caption = "Total Charges: Php " & FormatNumber(curReportsSubTotal + (curReportsSubTotal * 0.12), 2, True, True, True)
  ' Set the default value for our Cash Payment input field
  ' 0.00 signify that it need to be a currency
  txtCashPayment.Text = "0.00"
  
  If Dir(strPath & "\BasicPOS\SalesReports\") = "" Then
     MkDir (strPath & "\BasicPOS")
     MkDir (strPath & "\BasicPOS\SalesReports")
  End If
    
  ' I just used this as holder for the string path of reports
  lblSalesReportsHolder.Caption = strPath & "\BasicPOS\SalesReports"
  
End Sub


Private Sub txtCashPayment_Change()
 
 ' Validate that only numeric value is entered in our Cash Payment fields
 If Not IsNumeric(txtCashPayment) Then
    MsgBox "Please Enter a Numeric Value", vbInformation, "Numeric Value Check Failed"
    txtCashPayment.Text = "0.00"
 End If
 
End Sub

Private Sub txtCashPayment_LostFocus()
   
   ' We will format the number entered so it will look more as Currency
   If Val(txtCashPayment.Text) <> 0 Then
     txtCashPayment.Text = FormatNumber(txtCashPayment.Text, 2, vbTrue, vbTrue, vbTrue)
   End If
   
End Sub
