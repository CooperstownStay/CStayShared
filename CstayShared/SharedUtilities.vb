Imports System.Net
Imports System.Configuration.ConfigurationSettings
Imports System.Linq
Imports System.Data.Linq


Public Class SharedUtilities

  Inherits System.Web.UI.Page

  Public sSMTPServer As String = ""
  Public sSMTPUser As String = ""
  Public sSMTPPass As String = ""
  Public sConnString As String = ""
  Public sSystemEmailAddress As String = ""
  Public sErrorEmailAddress As String = ""
  Public sSendAllEmailAddress As String = ""
  Public sErrorRedirectPage As String = ""
  Public sAppEnvironment As String = ""
  Public sAppType As String = ""
  Public sMainWebsite As String = ""
  Public sSecureWebsite As String = ""
  Public bLiveTesting As Boolean = False
  Public Shared sDateForTesting As String = ""

  Public oErr As ErrorHandler
  Public oAuth As CStayAuthorizeNET
  Public oMyUtil As MyUtilities

  Public Sub New()

    oMyUtil = New MyUtilities(Me)
    oErr = New ErrorHandler(Me)
    oAuth = New CStayAuthorizeNET(oErr, Me)
    sSendAllEmailAddress = GetAppSetting("SendAllEmailsTo")

  End Sub

  Public Sub New(ByVal ConnString As String, ByVal SMTPServer As String, ByVal SMTPUser As String, ByVal SMTPPassword As String,
   ByVal SystemEmailAddress As String, ByVal ErrorEmailAddress As String, ByVal AppType As String, ByVal AppEnvironment As String,
   ByVal MainWebsite As String, ByVal SecureWebsite As String, ByVal LiveTesting As Boolean, ByVal DateForTesting As String, ByVal ErrorRedirectPage As String)

    sConnString = ConnString
    sSMTPServer = SMTPServer
    sSMTPUser = SMTPUser
    sSMTPPass = SMTPPassword
    sSystemEmailAddress = SystemEmailAddress
    sErrorEmailAddress = ErrorEmailAddress
    sAppEnvironment = AppEnvironment
    sAppType = AppType
    sMainWebsite = MainWebsite
    sSecureWebsite = SecureWebsite
    bLiveTesting = LiveTesting
    sDateForTesting = DateForTesting

    oMyUtil = New MyUtilities(Me)
    oErr = New ErrorHandler(Me, True)
    oAuth = New CStayAuthorizeNET(oErr, Me)

  End Sub


  Protected Overrides Sub Finalize()

    oMyUtil = Nothing
    oErr = Nothing
    oAuth = Nothing
    MyBase.Finalize()

  End Sub

#Region "Reservation and Payment Utilities"

  Public Sub GetBalanceDueDate(ByVal dArriveDate As Date, ByRef dBalanceDueDate As Date, ByRef dAbsStopDate As Date)

    Dim dNow As Date = GetDate()
    Dim dShorterDepositDueStartDate As Date = CDate("2/1/" & dArriveDate.Year.ToString)
    '		Dim dShorterDepositDueStartDate2 As Date = CDate("3/1/" & dArriveDate.Year.ToString)
    Dim dShorterDepositDueEndDate As Date = CDate("9/10/" & dNow.Year.ToString)
    dAbsStopDate = dNow.AddMonths(3)
    Dim dAbsStopDate2 As Date = DateAdd("d", -30, dArriveDate)

    If dShorterDepositDueStartDate < dNow Then
      If dNow > dAbsStopDate2 Then
        dBalanceDueDate = dNow.ToShortDateString
      Else
        If dAbsStopDate > dAbsStopDate2 Then
          dBalanceDueDate = dAbsStopDate2.ToShortDateString
        Else
          dBalanceDueDate = dAbsStopDate.ToShortDateString
        End If
      End If
    Else
      dBalanceDueDate = dNow.AddMonths(4).ToShortDateString
    End If
    If dBalanceDueDate > dAbsStopDate2 Then dBalanceDueDate = dAbsStopDate2.ToShortDateString


  End Sub

  Public Function AvailablePropertiesWhereSQL(ByVal dArriveDate As Date, ByVal dDepartDate As Date) As String
    AvailablePropertiesWhereSQL = " ((" &
     "ArriveDate    > convert(smalldatetime,'" & FormatDateTime(dArriveDate, vbShortDate) & "')  " &
     "and ArriveDate  < convert(smalldatetime,'" & FormatDateTime(dDepartDate, vbShortDate) & "'))  " &
     "or (ArriveDate = convert(smalldatetime,'" & FormatDateTime(dArriveDate, vbShortDate) & "'))  " &
     "or (DepartDate  = convert(smalldatetime,'" & FormatDateTime(dDepartDate, vbShortDate) & "'))  " &
     "or (ArriveDate  > convert(smalldatetime,'" & FormatDateTime(dArriveDate, vbShortDate) & "')  " &
     "and DepartDate  < convert(smalldatetime,'" & FormatDateTime(dDepartDate, vbShortDate) & "'))) "

  End Function


  Public Function CreateCancellationPolicy(ByVal dRentalFee As Double, Optional ByVal oCode As TableCodeData = Nothing) As String

    Dim sOut As String = "", sCancelFee1 As String = "", sCancelFee2 As String = ""

    CreateCancellationPolicy = ""
    sOut = GetCodeValue("Cancellation Policy", True, oCode)
    sCancelFee1 = FormatCurrency(dRentalFee * 0.1, 0)
    sCancelFee2 = FormatCurrency(dRentalFee * 0.1, 0)
    sOut = sOut.Replace("$CancellationFee$", FormatCurrency(sCancelFee1, 2))
    CreateCancellationPolicy = sOut

  End Function



  'Public Function CreateCIMPaymentInformationVB6(ByRef sResult As String, ByVal sGuestID As String, ByVal sGuestName As String,
  ' ByVal sGuestEmail As String, ByRef sCustomerCIMID As String, ByRef sPaymentCIMID() As String, ByRef sPaymentID() As String, ByRef sPaymentCategory() As String,
  ' ByRef sCCNumber() As String, ByRef sCCMonth() As String, ByRef sCCYear() As String, ByRef sCCFirstName() As String, ByRef sCCLastName() As String,
  ' ByRef sCCAddress() As String, ByRef sCCCity() As String, ByRef sCCState() As String, ByRef sCCZip() As String,
  ' Optional ByVal bTestMode As Boolean = False) As Boolean

  '  CreateCIMPaymentInformationVB6 = False
  '  Dim oAuth As New CStayAuthorizeNET(Me.oErr, Me)
  '  Try
  '    CreateCIMPaymentInformationVB6 = oAuth.CreateCIMPaymentInformation(sResult, sGuestID, sGuestName,
  '     sGuestEmail, sCustomerCIMID, sPaymentCIMID, sPaymentID, sPaymentCategory, sCCNumber, sCCMonth, sCCYear, sCCFirstName,
  '     sCCLastName, sCCAddress, sCCCity, sCCState, sCCZip, bTestMode)

  '  Catch ex As Exception
  '    sResult = oErr.GetError(ex)
  '  End Try
  '  oAuth = Nothing


  'End Function




  'Public Function GetCIMPaymentInformationVB6(ByRef sResult As String, ByRef sCustomerCIMID As String, ByRef sPaymentCIMID As String, ByRef sCCPaymentID() As String, ByRef sCCNumber() As String,
  'ByRef sCCFirstName() As String, ByRef sCCLastName() As String,
  'ByRef sCCAddress() As String, ByRef sCCCity() As String, ByRef sCCState() As String, ByRef sCCZip() As String,
  'ByVal bTestMode As Boolean, ByVal bGetAllPaymentsIfRequestedPaymentMissing As Boolean) As Boolean

  '  GetCIMPaymentInformationVB6 = False
  '  Dim oAuth As New CStayAuthorizeNET(Me.oErr, Me)
  '  Try
  '    GetCIMPaymentInformationVB6 = oAuth.GetCIMPaymentInformation(sResult, sCustomerCIMID, sPaymentCIMID, sCCPaymentID, sCCNumber,
  '    sCCFirstName, sCCLastName, sCCAddress, sCCCity, sCCState, sCCZip, bTestMode, bGetAllPaymentsIfRequestedPaymentMissing)

  '  Catch ex As Exception
  '    sResult = oErr.GetError(ex)
  '  End Try
  '  oAuth = Nothing



  'End Function



  Public Function SavePaymentVB6(ByRef sResult As String, ByRef sAuthNETReason As String, ByRef sAuthNET_AVSText As String,
   ByRef sCustomerCIMID As String, ByRef sPaymentCIMID As String, ByVal iBookingID As Integer, ByRef iPaymentID As Integer, ByVal sPaymentOrigin As String,
   ByVal sPaymentLocation As String, ByVal sEmailType As String, ByVal bChargeInTestMode As Boolean,
   ByVal dPaymentAmount As Double, ByVal dChargeAmount As Double, ByVal sPaymentCategory As String, ByVal sPaymentMethod As String,
   ByVal sPaymentEmail As String, ByRef iContactLogID As Integer, ByVal bFinalBalPayMade As Boolean, ByVal bFinalDepPayMade As Boolean,
   ByVal sEncryptPassword As String, ByVal sCCType As String, ByVal sCCNumber As String, ByVal sCCFirstName As String, ByVal sCCLastName As String,
   ByVal sCCExpMonth As String, ByVal sCCExpYear As String, ByVal sCCVerification As String, ByVal sCCAddress As String, ByVal sCCCity As String,
   ByVal sCCState As String, ByVal sCCZip As String, ByVal sBankName As String, ByVal sBankAccountName As String, ByVal sBankAccountNumber As String,
   ByVal sBankABANumber As String, ByVal sBankAccountType As String, ByVal bSkipMakingCharge As Boolean, ByVal bPutErrorInResult As Boolean,
   ByVal sCreditType As String, ByVal sPrevTransactionID As String, ByVal iPrevPaymentID As Integer, ByVal sPrevPaymentCategory As String,
  ByRef sAuthorizationTypeUsed As String, ByVal sUserName As String, ByVal bUseCIMGateway As Boolean, ByRef bCIMTransactionDeclined As Boolean) As Boolean

    SavePaymentVB6 = False

    Dim oBook As New TableBookings(sConnString, True), oGuest As New TableGuests(oBook.Transaction), oProp As New TableProperties(oBook.Transaction)
    Dim oHost As New TableHosts(oBook.Transaction), oCode As New TableCodeData(oBook.Transaction), oPay As Object = Nothing, bUseBookingAmounts As Boolean = True
    Dim oAuth As New CStayAuthorizeNET(Me.oErr, Me), oHostPay As New TableHostPayments(oBook.Transaction), oPrevPayment As Object = Nothing
    Dim oCIMGate As New TableCIMGatewayActivity(oBook.Transaction), sUserResult As String = ""

    If sPrevPaymentCategory.ToString = "Balance Due" Then
      oPrevPayment = New TableBalancePayments(oBook.Transaction)
    ElseIf sPrevPaymentCategory.ToString = "Deposit" Then
      oPrevPayment = New TableDepositPayments(oBook.Transaction)
    End If

    Try

      oGuest.Guest_ID_PK__Integer = oBook.Guest_ID__Integer
      oGuest.SelectData()
      oHost.Host_ID_PK__Integer = oBook.Host_ID__Integer
      oHost.SelectData()
      oProp.Property_ID_PK__Integer = oBook.Property_ID__Integer
      oProp.SelectData()


      SavePaymentVB6 = SavePayment(sResult, sUserResult, oBook, oPay, oGuest, oProp, oHost, oHostPay, oCIMGate, sCustomerCIMID, sPaymentCIMID, iBookingID, iPaymentID,
       sPaymentOrigin, sPaymentLocation, sEmailType, False, bChargeInTestMode,
       dPaymentAmount, dChargeAmount, sPaymentCategory, sPaymentMethod,
       sPaymentEmail, iContactLogID, bFinalBalPayMade, bFinalDepPayMade,
       oAuth, sEncryptPassword, sCCType, sCCNumber, sCCFirstName, sCCLastName,
       sCCExpMonth, sCCExpYear, sCCVerification, sCCAddress, sCCCity,
       sCCState, sCCZip, sBankName, sBankAccountName, sBankAccountNumber, sBankABANumber, sBankAccountType, bSkipMakingCharge, bPutErrorInResult, sCreditType,
       sPrevTransactionID, iPrevPaymentID, sPrevPaymentCategory, oPrevPayment, sAuthorizationTypeUsed, sUserName, , bUseCIMGateway, sConnString, bCIMTransactionDeclined)

      sAuthNETReason = oAuth.RespReasonText
      sAuthNET_AVSText = oAuth.RespAVSText

    Catch ex As Exception
      sResult = oErr.GetError(ex)
    End Try
    oBook = Nothing
    oGuest = Nothing
    oHost = Nothing
    oProp = Nothing
    oPay = Nothing
    oCode = Nothing
    oHostPay = Nothing
    oAuth = Nothing
    oPrevPayment = Nothing
  End Function



  Public Function CreateEmailHTMLVB6(ByRef sResult As String, ByRef sEmail As String, ByVal sType As String, ByVal sPaymentCategory As String, ByVal sEmailType As String,
  ByVal sBookingID As String, ByVal sPaymentID As String, ByVal sCCamount As String) As Boolean

    CreateEmailHTMLVB6 = False
    Dim oBook As New TableBookings(sConnString), oGuest As New TableGuests(sConnString), oProp As New TableProperties(sConnString)
    Dim oHost As New TableHosts(sConnString), oCode As New TableCodeData(sConnString), oPay As Object = Nothing, bUseBookingAmounts As Boolean = True

    Try


      oBook.Booking_ID_PK__Integer = CNullI(sBookingID)
      oBook.SelectData()
      If oBook.Booking_ID_PK__Integer > 0 Then

        If CNullD(sPaymentID) > 0 Then
          If sPaymentCategory = "Deposit" Then
            oPay = New TableDepositPayments(sConnString)
          ElseIf sPaymentCategory Like "Balance*" Then
            oPay = New TableBalancePayments(sConnString)
          End If
          oPay.PaymentID_PK__Integer = CNullI(sPaymentID)
          oPay.SelectData()
          If oPay.PaymentID_PK__Integer = 0 Then
            oPay = Nothing
          Else
            bUseBookingAmounts = False
          End If
        Else
          ' logic to choose $PaymentAmount$ in absence of a payment record
          If sPaymentCategory = "Deposit" Then
            sCCamount = oBook.DepositAmount__Numeric
          ElseIf sPaymentCategory Like "Balance*" Then
            sCCamount = oBook.BalanceAmount__Numeric
          End If
          If oBook.Status__String = "Booked, Waiting for First Payment" Then bUseBookingAmounts = False
        End If

        oGuest.Guest_ID_PK__Integer = oBook.Guest_ID__Integer
        oGuest.SelectData()
        oHost.Host_ID_PK__Integer = oBook.Host_ID__Integer
        oHost.SelectData()
        oProp.Property_ID_PK__Integer = oBook.Property_ID__Integer
        oProp.SelectData()

        Dim sImageURL As String = "", sTemp As String = ""
        CreateEmailHTMLVB6 = CreateEmailHTML(sResult, sType, sPaymentCategory, sEmailType, sEmail, sTemp, CNullI(sBookingID), sImageURL,
         oPay, oBook, oGuest, oProp, oHost, oCode, "Telephone", , , , , , , sCCamount, bUseBookingAmounts, True)
      End If
    Catch ex As Exception
      sResult = oErr.GetError(ex)
    End Try
    oBook = Nothing
    oHost = Nothing
    oProp = Nothing
    oPay = Nothing
    oCode = Nothing
  End Function


  Public Function CreateReservationSummaryVB6(ByRef sResult As String, ByVal sBookingID As String, ByVal sRemainingAmountMessage As String) As Boolean

    CreateReservationSummaryVB6 = False
    Dim oBook As New TableBookings(sConnString), oGuest As New TableGuests(sConnString), oProp As New TableProperties(sConnString)
    Dim oHost As New TableHosts(sConnString), oCode As New TableCodeData(sConnString)
    Try

      Dim dRemainingBalance As Double = 0, sGuestName As String = "", dDepositRemaining As Double = 0, dBalanceRemaining As Double = 0
      Dim dDepositPayments As Double = 0, dBalancePayments As Double = 0, bUseBookingAmounts As Boolean = False

      Dim sAdminEmail As String = "", sEmail As String = ""

      oBook.Booking_ID_PK__Integer = CNullI(sBookingID)
      oBook.SelectData()
      If oBook.Booking_ID_PK__Integer > 0 Then

        oGuest.Guest_ID_PK__Integer = oBook.Guest_ID__Integer
        oGuest.SelectData()
        oHost.Host_ID_PK__Integer = oBook.Host_ID__Integer
        oHost.SelectData()
        oProp.Property_ID_PK__Integer = oBook.Property_ID__Integer
        oProp.SelectData()

        CreateReservationSummaryVB6 = CreateReservationSummary(sResult, True, oBook, oGuest, oProp, dRemainingBalance,
         sGuestName, dDepositRemaining, dBalanceRemaining, dDepositPayments, dBalancePayments, CBool(sRemainingAmountMessage), oCode, bUseBookingAmounts)
      End If
    Catch ex As Exception
      sResult = oErr.GetError(ex)
    End Try
    oBook = Nothing
    oHost = Nothing
    oProp = Nothing
    oCode = Nothing

  End Function



  Public Function CreateReservationInfoVB6(ByRef sResult As String, ByVal sBookingID As String, bIncludeCheckInDetails As String) As String

    CreateReservationInfoVB6 = False
    Dim oBook As New TableBookings(sConnString), oGuest As New TableGuests(sConnString), oProp As New TableProperties(sConnString)
    Dim oHost As New TableHosts(sConnString), oCode As New TableCodeData(sConnString)

    Try

      Dim dRemainingBalance As Double = 0, sGuestName As String = "", dDepositRemaining As Double = 0, dBalanceRemaining As Double = 0
      Dim dDepositPayments As Double = 0, dBalancePayments As Double = 0, sImageURL As String = ""

      Dim sAdminEmail As String = "", sEmail As String = ""

      oBook.Booking_ID_PK__Integer = CNullI(sBookingID)
      oBook.SelectData()
      If oBook.Booking_ID_PK__Integer > 0 Then

        oGuest.Guest_ID_PK__Integer = oBook.Guest_ID__Integer
        oGuest.SelectData()
        oHost.Host_ID_PK__Integer = oBook.Host_ID__Integer
        oHost.SelectData()
        oProp.Property_ID_PK__Integer = oBook.Property_ID__Integer
        oProp.SelectData()

        CreateReservationInfoVB6 = CreateReservationInfo(sResult, oBook, oGuest, oProp, oHost,
         sGuestName, oCode, sImageURL, True, CBool(bIncludeCheckInDetails))

      End If
    Catch ex As Exception
      sResult = "ERROR:" & vbCrLf & "Message: " & ex.Message.ToString & vbCrLf & "Site:  " & ex.TargetSite.ToString & vbCrLf & "Stack Trace:  " & ex.StackTrace.ToString
      'sResult = oErr.GetError(ex)
    End Try
    oBook = Nothing
    oHost = Nothing
    oProp = Nothing
    oCode = Nothing

  End Function



  Public Function CreateEmailHTML(ByRef sResult As String, ByVal sType As String, ByVal sPaymentCategory As String, ByVal sEmailType As String, ByRef sEmail As String,
   ByRef sEmailAdminVersion As String, ByVal iBookingID As Integer, ByVal sImageURL As String,
   ByRef oPay As Object, ByRef oBook As TableBookings, ByRef oGuest As TableGuests,
   ByRef oProp As TableProperties, ByRef oHost As TableHosts,
   Optional ByRef oCode As TableCodeData = Nothing, Optional ByVal sLocationMade As String = "Online",
   Optional ByVal sCCNumber As String = "", Optional ByVal sCCName As String = "",
   Optional ByVal sCCAddress As String = "", Optional ByVal sCCCity As String = "",
   Optional ByVal sCCState As String = "", Optional ByVal sCCZip As String = "", Optional ByRef sCCamount As String = "",
   Optional ByVal bUseBookingAmounts As Boolean = False, Optional ByVal bPutErrorInResult As Boolean = False) As Boolean

    Dim sGuestName As String = "", sPaymentTag As String = "", dTotalRemaining As Double = 0
    Dim sResrvOut As String = "", sPolicyOut As String = "", sGuestOut As String = "", sAccomOut As String = "", sFullAmount As String = ""
    Dim sReservInfo As String = "", sRemainingAmountMessage As String = "", sTemp As String = ""
    Dim sCancelFee1 As String = "", dRentalTotal As Double = 0, sReservAgree As String = "", sReservationSummary As String = ""
    Dim dDepositRemaining As Double = 0, dBalanceRemaining As Double = 0, sPropLink As String = "", dCCAmount As Double = 0
    Dim dDepositPayments As Double = 0, dBalancePayments As Double = 0, sTitle As String = "", sCancelPolicy As String = "", sThisResult As String = ""
    Dim sStyleSheet As String = GetCodeValue("ReservationEmailStyleSheet", False, oCode), sEmailText As String = "", sBankAmount As String = ""
    Dim sBankName As String = "", sBankAccountName As String = "", sBankAccountNumber As String = "", sBankABANumber As String = "", sBankAccountType As String = "", sPaymentOptions As String = ""
    Dim sAdminReservInfo As String = ""


    CreateEmailHTML = False
    Try

      If sEmailType = "" Then sEmailType = "Default HTML Body"
      sEmail = GetCodeValue(sEmailType, True, oCode, False)
      If sEmail <> "" Then

        sEmailAdminVersion = GetCodeValue("Reservation Email Admin Version", True, oCode, False)
        If Not CreateReservationInfo(sThisResult, oBook, oGuest, oProp, oHost, sGuestName, oCode, sImageURL) Then Throw New Exception(sThisResult)
        sReservInfo = sThisResult
        If Not CreateReservationInfo(sThisResult, oBook, oGuest, oProp, oHost, sGuestName, oCode, sImageURL, True, , True) Then Throw New Exception(sThisResult)
        sAdminReservInfo = sThisResult


        sReservAgree = GetCodeValue("Reservation Agreement", True, oCode)
        If sLocationMade <> "Online" Then
          sPaymentOptions = GetCodeValue("Payment Options", True, oCode)
          sPaymentOptions = Replace(sPaymentOptions, "$ReservationNumber$", iBookingID.ToString)
        End If


        If Not CreateReservationSummary(sThisResult, True, oBook, oGuest, oProp, dTotalRemaining, sGuestName,
         dDepositRemaining, dBalanceRemaining, dDepositPayments, dBalancePayments, (sType = "Reservation"), oCode, bUseBookingAmounts) Then Throw New Exception(sThisResult)
        sReservationSummary = sThisResult

        dRentalTotal = oBook.Rate__Numeric + oBook.TaxDue__Numeric
        sCancelFee1 = FormatCurrency(dRentalTotal * 0.1, 0)

        If sType = "Reservation" Or sType = "Payment" Then
          If sType = "Reservation" Then
            If dTotalRemaining > 0 Then
              sPaymentTag = "first "
            Else
              sPaymentTag = "full amount"
            End If
          Else
            If dTotalRemaining > 0 Then
              sPaymentTag = "partial "
            Else
              sPaymentTag = "full balance"
            End If
          End If
        End If

        If Not oPay Is Nothing Then
          dCCAmount = oPay.Amount__Numeric
          If oPay.Type__String = "Card" Then
            sCCNumber = "..." & oPay.CCNumber__String
            sCCName = oPay.CCName__String
            sCCAddress = oPay.CCAddress__String
            sCCCity = oPay.CCCity__String
            sCCState = oPay.CCState__String
            sCCZip = oPay.CCZip__String
          ElseIf oPay.Type__String = "eCheck" Then
            sBankABANumber = oPay.BankRoutingNumber__String
            sBankAccountName = oPay.BankAccountName__String
            sBankAccountNumber = "..." & oPay.BankAccountNumber__String
            sBankAccountType = oPay.BankAccountType__String
            sBankName = oPay.BankName__String
          End If
          sCCamount = FormatCurrency(dCCAmount, 2)
          sBankAmount = sCCamount
        Else
          If IsNumeric(sCCamount) Then
            dCCAmount = CNullD(sCCamount)
            sCCamount = FormatCurrency(dCCAmount, 2)
          End If
        End If
        sCancelPolicy = CreateCancellationPolicy(dRentalTotal)


        If sCCName <> "" Then
          sEmail = ClipHTMLSectionToUse(sEmail, "<echeck>", True)
          sEmail = sEmail.Replace("$CreditCardName$", sCCName)
          sEmail = sEmail.Replace("$CreditCardAddress$", sCCAddress)
          sEmail = sEmail.Replace("$CreditCardNumber$", sCCNumber)
          sEmail = sEmail.Replace("$CreditCardCity$", sCCCity)
          sEmail = sEmail.Replace("$CreditCardState$", sCCState)
          sEmail = sEmail.Replace("$CreditCardZip$", sCCZip)
          sEmailAdminVersion = ClipHTMLSectionToUse(sEmailAdminVersion, "<echeck>", True)
          sEmailAdminVersion = sEmailAdminVersion.Replace("$CreditCardName$", sCCName)
          sEmailAdminVersion = sEmailAdminVersion.Replace("$CreditCardAddress$", sCCAddress)
          sEmailAdminVersion = sEmailAdminVersion.Replace("$CreditCardNumber$", sCCNumber)
          sEmailAdminVersion = sEmailAdminVersion.Replace("$CreditCardCity$", sCCCity)
          sEmailAdminVersion = sEmailAdminVersion.Replace("$CreditCardState$", sCCState)
          sEmailAdminVersion = sEmailAdminVersion.Replace("$CreditCardZip$", sCCZip)
        Else
          sEmail = ClipHTMLSectionToUse(sEmail, "<credit_card>", True)
          sEmailAdminVersion = ClipHTMLSectionToUse(sEmailAdminVersion, "<credit_card>", True)
          If sBankName.ToString <> "" Then
            sEmail = sEmail.Replace("$BankPaymentAmount$", sBankAmount)
            sEmail = sEmail.Replace("$BankName$", sBankName)
            sEmail = sEmail.Replace("$BankABANumber$", sBankABANumber)
            sEmail = sEmail.Replace("$BankAccountName$", sBankAccountName)
            sEmail = sEmail.Replace("$BankAccountNumber$", sBankAccountNumber)
            sEmail = sEmail.Replace("$BankAccountType$", sBankAccountType)
            sEmailAdminVersion = sEmailAdminVersion.Replace("$BankPaymentAmount$", sBankAmount)
            sEmailAdminVersion = sEmailAdminVersion.Replace("$BankName$", sBankName)
            sEmailAdminVersion = sEmailAdminVersion.Replace("$BankABANumber$", sBankABANumber)
            sEmailAdminVersion = sEmailAdminVersion.Replace("$BankAccountName$", sBankAccountName)
            sEmailAdminVersion = sEmailAdminVersion.Replace("$BankAccountNumber$", sBankAccountNumber)
            sEmailAdminVersion = sEmailAdminVersion.Replace("$BankAccountType$", sBankAccountType)
          End If
        End If

        sEmail = sEmail.Replace("$CreditCardAmount$", sCCamount)
        sEmail = sEmail.Replace("$PaymentLocation$ ", "")
        sEmail = sEmail.Replace("$PaymentType$", sPaymentTag)
        sEmail = sEmail.Replace("$PaymentAmount$", sCCamount)
        sEmail = sEmail.Replace("$$PAGE_TITLE$$", sTitle)

        sEmailAdminVersion = sEmailAdminVersion.Replace("$CreditCardAmount$", sCCamount)
        sEmailAdminVersion = sEmailAdminVersion.Replace("$PaymentLocation$ ", "")
        sEmailAdminVersion = sEmailAdminVersion.Replace("$PaymentType$", sPaymentTag)
        sEmailAdminVersion = sEmailAdminVersion.Replace("$PaymentAmount$", sCCamount)
        sEmailAdminVersion = sEmailAdminVersion.Replace("$$PAGE_TITLE$$", sTitle)
        sEmailAdminVersion = sEmailAdminVersion.Replace("$EmailType$", sType)
        sEmailAdminVersion = sEmailAdminVersion.Replace("background-color: rgb(238, 238, 187)", "")


        sEmail = Replace(sEmail, "$$RESERVATION_INFO$$", sReservInfo)
        sEmail = Replace(sEmail, "$$RESERVATION_SUMMARY$$", sReservationSummary)
        sEmail = Replace(sEmail, "$$PAYMENT_OPTIONS$$", sPaymentOptions)
        sEmail = Replace(sEmail, "$$CANCELLATION_POLICY$$", sCancelPolicy)
        sEmail = Replace(sEmail, "$$RESERVATION_AGREEMENT$$", sReservAgree)
        sEmail = Replace(sEmail, "background-color: rgb(238, 238, 187)", "")

        sEmailAdminVersion = Replace(sEmailAdminVersion, "$$RESERVATION_INFO$$", sAdminReservInfo)
        sEmailAdminVersion = Replace(sEmailAdminVersion, "$$RESERVATION_SUMMARY$$", sReservationSummary)
        sEmailAdminVersion = Replace(sEmailAdminVersion, "$$PAYMENT_OPTIONS$$", sPaymentOptions)
        sEmailAdminVersion = Replace(sEmailAdminVersion, "$$CANCELLATION_POLICY$$", sCancelPolicy)
        sEmailAdminVersion = Replace(sEmailAdminVersion, "$$RESERVATION_AGREEMENT$$", sReservAgree)
        sEmailAdminVersion = ClipHTMLSectionToUse(sEmailAdminVersion, "<damage_deposit>", True)


        If sStyleSheet.ToString <> "" Then sEmail = ClipHTMLSectionToUse(sEmail, "<style_sheet>", True, sStyleSheet)
        If sStyleSheet.ToString <> "" Then sEmailAdminVersion = ClipHTMLSectionToUse(sEmailAdminVersion, "<style_sheet>", True, sStyleSheet)

        CreateEmailHTML = True
      End If

    Catch ex As Exception
      If bPutErrorInResult Then sResult = "ERROR:" & vbCrLf & "Message: " & ex.Message.ToString & vbCrLf & "Site:  " & ex.TargetSite.ToString & vbCrLf & "Stack Trace:  " & ex.StackTrace.ToString
      oErr.HandleError(ex, False)
    End Try

  End Function


  Public Function RemoveEmailTags(ByVal sMsgIn As String) As String

    RemoveEmailTags = sMsgIn
    RemoveEmailTags = RemoveEmailTags.Replace("<ADMIN_EXCLUDE>", "")
    RemoveEmailTags = RemoveEmailTags.Replace("<ADMIN_INCLUDE>", "")
    RemoveEmailTags = RemoveEmailTags.Replace("<CHECK_IN_DETAILS>", "")
    RemoveEmailTags = RemoveEmailTags.Replace("<MAIN_BODY>", "")
    RemoveEmailTags = RemoveEmailTags.Replace("<FOOTER>", "")
    RemoveEmailTags = RemoveEmailTags.Replace("<STYLE_SHEET>", "")
    RemoveEmailTags = RemoveEmailTags.Replace("<CREDIT_CARD>", "")
    RemoveEmailTags = RemoveEmailTags.Replace("<ECHECK>", "")
    RemoveEmailTags = RemoveEmailTags.Replace("<SUMMARY_TITLE>", "")
    RemoveEmailTags = RemoveEmailTags.Replace("<FIRST_PAYMENT_SECTION>", "")
    RemoveEmailTags = RemoveEmailTags.Replace("<REMAINING_AMOUNT_SECTION>", "")

  End Function


  Public Function CreateReservationInfo(ByRef sResult As String, ByRef oBook As TableBookings, ByRef oGuest As TableGuests, ByRef oProp As TableProperties, ByRef oHost As TableHosts,
    ByVal sGuestName As String, ByRef oCode As TableCodeData, ByVal sImageURL As String, Optional ByVal bPutErrorInResult As Boolean = False, Optional bIncludeCheckInDetails As Boolean = True, Optional bAdminVersion As Boolean = False) As Boolean

    CreateReservationInfo = False
    Dim sData() As String, sHostName As String = "", sReservInfo As String = "", sPropLink As String, sCheckInDetails As String = "", sGPS As String = "", sAreaMapLink = ""
    Dim oHouse As HouseRentalDataContext = MyData.GetHouseRentalContext(False, sConnString), iProperty_ID As Integer = oBook.Property_ID__Integer
    Try

      If bAdminVersion Then
        sReservInfo = GetCodeValue("Reservation Information Admin Version", True, oCode)
      Else
        sReservInfo = GetCodeValue("Reservation Information", True, oCode)
      End If
      If bIncludeCheckInDetails Then
        sCheckInDetails = "<div style='margin: 10px; width: 440px; padding-left: 15px; border:2px solid #990000;'>" & GetCodeValue("Reservation Check-In Details", True, oCode) & "</div>"
        sReservInfo = sReservInfo.Replace("<CHECK_IN_DETAILS>", "<CHECK_IN_DETAILS>" & sCheckInDetails & "<CHECK_IN_DETAILS>")
        sReservInfo = sReservInfo.Replace("<check_in_details>", "<check_in_details>" & sCheckInDetails & "<check_in_details>")
      End If


      If sReservInfo <> "" Then
        ' BJR kluge until start storing first/last name in sepearte fields
        sHostName = oHost.HostName__String
        sData = Split(sHostName, ",")
        If UBound(sData) > 0 Then
          sHostName = sData(1).Trim & " " & sData(0)
        End If

        If sGuestName.ToString = "" Then
          sGuestName = oGuest.GuestName__String.ToString
          sData = Split(sGuestName, ",")
          If UBound(sData) > 0 Then
            sGuestName = sData(1).Trim & " " & sData(0)
          End If
        End If

        Dim oPropGroup As vwPropertiesWithGroup = Nothing, iGroupID As Integer = 0
          Dim iImagePropID As Integer = iProperty_ID, sImageFile As String = ""

        oPropGroup = (From x In oHouse.vwPropertiesWithGroups Where x.Property_ID = iProperty_ID).FirstOrDefault
        If oPropGroup IsNot Nothing Then
          If CNullI(oPropGroup.GroupID) > 0 Then
            iGroupID = oPropGroup.GroupID
            oPropGroup = (From x In oHouse.vwPropertiesWithGroups Where x.GroupID = iGroupID Order By x.SequenceNumber).FirstOrDefault
            End If
          End If

        sAreaMapLink = sMainWebsite & "/AreaMap.aspx?Props=" & iProperty_ID & ":"
        If iGroupID > 0 Then sAreaMapLink &= iGroupID

        If sImageURL = "" Then
          If oPropGroup IsNot Nothing Then
            sImageFile = oPropGroup.MasterImageFile
            iImagePropID = oPropGroup.Property_ID
          End If
          sImageURL = sMainWebsite & "/PropertyPhotos/" & iImagePropID & "/" & sImageFile
        End If

        If oProp.Latitude__Numeric <> 0 And oProp.Longitude__Numeric <> 0 Then
          sGPS = "<span style=""font-size:10px;"" >GPS: " & FormatNumber(oProp.Latitude__Numeric, 5) & ", " & FormatNumber(oProp.Longitude__Numeric, 5) & "</span>"

        End If

        sPropLink = sMainWebsite & "/" & MyData.SetPropertyName(oProp.PropertyName__String, oProp.Category_ID__Integer)
        sImageURL = "<a href=""" & sPropLink & """ target=""_blank""><img src=""" & sImageURL & """ border=""0"" alt=""""></a>"

        sReservInfo = sReservInfo.Replace("$PropertyName$", "<a href=""" & sPropLink & """ target=""_blank"">" & oProp.PropertyName__String & "</a>")
        sReservInfo = sReservInfo.Replace("$PropertyAddress$", oProp.Address__String)
        sReservInfo = sReservInfo.Replace("$PropertyCity$", oProp.City__String)
        sReservInfo = sReservInfo.Replace("$PropertyZip$", oProp.Zip__String)
        sReservInfo = sReservInfo.Replace("$PropertyPhone$", oProp.Phone__String)
        sReservInfo = sReservInfo.Replace("$PropertyGPS$", sGPS)
        sReservInfo = sReservInfo.Replace("$Sleeps$", oProp.Sleeps__Integer)
        sReservInfo = sReservInfo.Replace("$HostName$", sHostName)
        sReservInfo = sReservInfo.Replace("$HostAddress$", oHost.Address__String)
        sReservInfo = sReservInfo.Replace("$HostCity$", oHost.City__String)
        sReservInfo = sReservInfo.Replace("$HostZip$", oHost.Zip__String)
        sReservInfo = sReservInfo.Replace("$HostHomePhone$", oHost.HomePhone__String & " (home)")
        sReservInfo = sReservInfo.Replace("$HostWorkPhone$", IIf(oHost.WorkPhone__String = "", "", "<br>&nbsp;&nbsp;&nbsp;" & oHost.WorkPhone__String & " (work)"))
        sReservInfo = sReservInfo.Replace("$HostCellPhone$", IIf(oHost.CellPhone__String = "", "", "<br>&nbsp;&nbsp;&nbsp;" & oHost.CellPhone__String & " (cell)"))
        sReservInfo = sReservInfo.Replace("$HostCellPhone2$", IIf(oHost.CellPhone2__String = "", "", "<br>&nbsp;&nbsp;&nbsp;" & oHost.CellPhone2__String & " (cell 2)"))
        sReservInfo = sReservInfo.Replace("$HostEmail$", oHost.Email__String)
        sReservInfo = sReservInfo.Replace("$HostEmail2$", IIf(oHost.Email2__String = "", "", "<br>&nbsp;&nbsp;&nbsp;" & oHost.Email2__String))
        sReservInfo = sReservInfo.Replace("$GuestName$", sGuestName)
        sReservInfo = sReservInfo.Replace("$GuestAddress$", oGuest.Address__String)
        sReservInfo = sReservInfo.Replace("$GuestCity$", oGuest.City__String & ", " & oGuest.State__String)
        sReservInfo = sReservInfo.Replace("$GuestZip$", oGuest.Zip__String)
        sReservInfo = sReservInfo.Replace("$GuestHomePhone$", oGuest.HomePhone__String & " (home)")
        sReservInfo = sReservInfo.Replace("$GuestWorkPhone$", IIf(oGuest.WorkPhone__String = "", "", "<br>&nbsp;&nbsp;&nbsp;" & oGuest.WorkPhone__String & " (work)"))
        sReservInfo = sReservInfo.Replace("$GuestCellPhone$", IIf(oGuest.CellPhone__String = "", "", "<br>&nbsp;&nbsp;&nbsp;" & oGuest.CellPhone__String & " (cell)"))
        sReservInfo = sReservInfo.Replace("$GuestCellPhone2$", IIf(oGuest.CellPhone2__String = "", "", "<br>&nbsp;&nbsp;&nbsp;" & oGuest.CellPhone2__String & " (cell 2)"))
        sReservInfo = sReservInfo.Replace("$GuestEmail$", oGuest.Email__String)
        sReservInfo = sReservInfo.Replace("$GuestEmail2$", IIf(oGuest.Email2__String = "", "", "<br>&nbsp;&nbsp;&nbsp;" & oGuest.Email2__String))
        sReservInfo = sReservInfo.Replace("$Adults$", oBook.Adults__Integer)
        sReservInfo = sReservInfo.Replace("$Children$", oBook.Children__Integer)
        sReservInfo = sReservInfo.Replace("$Teens$", oBook.Teens__Integer)
        sReservInfo = sReservInfo.Replace("$PropertyWebImage$", sImageURL)
        sReservInfo = sReservInfo.Replace("$AreaMapLink$", sAreaMapLink)

        sResult = sReservInfo
        CreateReservationInfo = True
      End If
    Catch ex As Exception
      If bPutErrorInResult Then sResult = "ERROR:" & vbCrLf & "Message: " & ex.Message.ToString & vbCrLf & "Site:  " & ex.TargetSite.ToString & vbCrLf & "Stack Trace:  " & ex.StackTrace.ToString
      oErr.HandleError(ex, False)
    End Try
    MyData.DisposeHouseRentalContext(oHouse)
  End Function


  Public Function CreateReservationSummary(ByRef sResult As String, ByVal bCreatTitle As Boolean,
   ByRef oBook As TableBookings, ByRef oGuest As TableGuests, ByRef oProp As TableProperties, ByRef dTotalPaymentRemaining As Double,
   ByRef sGuestName As String, ByRef dFirstPaymentRemaining As Double, ByRef dBalancePaymentRemaining As Double,
   ByRef dFirstPaymentReceived As Double, ByRef dBalancePaymentReceived As Double, Optional ByRef bAutoChargeMessage As Boolean = False,
   Optional ByRef oCode As TableCodeData = Nothing, Optional ByRef bUseBookingAmounts As Boolean = False, Optional ByVal bPutErrorInResult As Boolean = False) As Boolean

    CreateReservationSummary = False
    Dim sTemp As String = ""


    Try

      Dim sTemplate As String = "", sData() As String, dArrivalDate As Date, dDepartureDate As Date, oDB As New DBUtilities
      Dim dTotalPaymentDue As Double = 0, sFirstPaymentDueDate As String = "", sBalancePaymentDueDate As String = ""
      Dim sRemainingAmountMessage As String = "", dTotalPaymentReceived As Double = 0, dFirstPaymentDue As Double = 0, dBalancePaymentDue As Double = 0
      Dim sFirstPaymentDue As String = "", sBalancePaymentDue As String = "", sFirstPaymentRemaining As String = "", sBalancePaymentRemaining As String = ""
      Dim dtFirstPaymentDueDate As Date, dtBalancePaymentDueDate As Date, bWaitingFirstPayment As Boolean = False
      Dim sTotalRemaining As String = "", sAutoChargePaymentType As String = "", dCreditCardFee As Double = 0

      ' BJR kluge until start storing first/last name in sepearte fields
      sGuestName = oGuest.GuestName__String.ToString
      sData = Split(sGuestName, ",")
      If UBound(sData) > 0 Then
        sGuestName = sData(1).Trim & " " & sData(0)
      End If

      dArrivalDate = CDate(oBook.ArriveDate__Date)
      dDepartureDate = CDate(oBook.DepartDate__Date)

      If bUseBookingAmounts Then
        dFirstPaymentReceived = oBook.DepositAmount__Numeric
        dBalancePaymentReceived = oBook.BalanceAmount__Numeric
      Else
        If oBook.Transaction Is Nothing Then
          oDB.SelectData("EXEC spGetBookingPayments " & oBook.Booking_ID_PK__Integer, , oBook.Connection)
        Else
          oDB.SelectData("EXEC spGetBookingPayments " & oBook.Booking_ID_PK__Integer, , oBook.Connection, oBook.Transaction)
        End If

        If oDB.MoveNext Then
          dFirstPaymentReceived = CNullD(oDB.CurrentRow("DepositPayments"))
          dBalancePaymentReceived = CNullD(oDB.CurrentRow("BalancePayments"))
        Else
          dFirstPaymentReceived = oBook.DepositAmountPaid__Numeric
          dBalancePaymentReceived = oBook.BalanceAmount__Numeric
        End If
      End If
      dTotalPaymentReceived = dFirstPaymentReceived + dBalancePaymentReceived

      If oBook.Status__String = "Booked, Waiting for First Payment" And dTotalPaymentReceived = 0 Then
        bWaitingFirstPayment = True
        sTemplate = GetCodeValue("Reservation Summary First Payment Due", True, oCode)
      Else
        sTemplate = GetCodeValue("Reservation Summary", True, oCode)
      End If


      dTotalPaymentDue = oBook.Rate__Numeric + oBook.TaxDue__Numeric
      dFirstPaymentDue = dTotalPaymentDue / 2
      dBalancePaymentDue = dFirstPaymentDue

      If oBook.Status__String = "Cancelled" Then
        dFirstPaymentRemaining = 0
        dBalancePaymentRemaining = 0
        dTotalPaymentRemaining = 0
      Else
        ' If all or most of balance paid in full up front
        If dFirstPaymentReceived > dFirstPaymentDue Then
          dBalancePaymentRemaining = 0
          dFirstPaymentRemaining = dTotalPaymentDue - dFirstPaymentReceived
        Else
          dFirstPaymentRemaining = dFirstPaymentDue - dFirstPaymentReceived
          dBalancePaymentRemaining = dBalancePaymentDue - dBalancePaymentReceived
        End If
        dTotalPaymentRemaining = dTotalPaymentDue - dTotalPaymentReceived
        If dTotalPaymentRemaining = 0 Then
          sTotalRemaining = "<b>Paid in Full</b>"
        Else
          sTotalRemaining = FormatCurrency(dTotalPaymentRemaining, 2)
        End If
      End If



      dtFirstPaymentDueDate = CDate(oBook.DepositDue__Date)
      dtBalancePaymentDueDate = CDate(oBook.BalanceDueDate__Date)

      sFirstPaymentDueDate = "&nbsp;&nbsp;<b>(due by " & Format(dtFirstPaymentDueDate, "MMMM ") & " " & dtFirstPaymentDueDate.Day & ", " & dtFirstPaymentDueDate.Year & ")</b>"
      sBalancePaymentDueDate = "&nbsp;&nbsp;<b>(due by " & Format(dtBalancePaymentDueDate, "MMMM ") & " " & dtBalancePaymentDueDate.Day & ", " & dtBalancePaymentDueDate.Year & ")</b>"

      If bAutoChargeMessage And dTotalPaymentRemaining > 0 Then

        If oBook.AutoCharge__Integer = 1 Then
          sAutoChargePaymentType = "credit card"

          sRemainingAmountMessage = "&nbsp;&nbsp; * This will be charged to your " & sAutoChargePaymentType.ToString & " on " & CDate(oBook.BalanceDueDate__Date).ToShortDateString
        End If

      End If

      If bWaitingFirstPayment Then
        sFirstPaymentDue = FormatCurrency(oBook.DepositAmount__Numeric, 2)
        sBalancePaymentDue = FormatCurrency(oBook.BalanceDue__Numeric, 2)
      Else
        sFirstPaymentDue = FormatCurrency(dFirstPaymentDue, 2)
        sBalancePaymentDue = FormatCurrency(dBalancePaymentDue, 2)

        If dFirstPaymentRemaining = 0 Then sFirstPaymentDueDate = ""
        If dBalancePaymentRemaining = 0 Then sBalancePaymentDueDate = ""
      End If


      If Not bCreatTitle Then
        sTemplate = ClipHTMLSectionToUse(sTemplate, "<SUMMARY_TITLE>", True)
        sTemplate = ClipHTMLSectionToUse(sTemplate, "<summary_title>", True)
      End If

      If sRemainingAmountMessage = "" Then
        sTemplate = ClipHTMLSectionToUse(sTemplate, "<remaining_amount_section>", True)
      Else
        If Not bWaitingFirstPayment Then
          sBalancePaymentDueDate = "*"
          sTemplate = sTemplate.Replace("$RemainingAmountMessage$", sRemainingAmountMessage)
        End If
      End If

      If dCreditCardFee = 0 Then
        sTemplate = ClipHTMLSectionToUse(sTemplate, "<credit_card_fee_section>", True)
      Else
        sTemplate = sTemplate.Replace("$CreditCardConvienceFee$", FormatCurrency(dCreditCardFee, 2))
      End If

      sTemplate = sTemplate.Replace("$GuestName$", sGuestName)
      sTemplate = sTemplate.Replace("$PropertyName$", oProp.PropertyName__String)
      sTemplate = sTemplate.Replace("$ReservationNumber$", oBook.Booking_ID_PK__Integer)
      sTemplate = sTemplate.Replace("$ArriveDate$", "<i>" & Format(dArrivalDate, "dddd, MMMM ") & dArrivalDate.Day & ", " & dArrivalDate.Year & "</i>")
      sTemplate = sTemplate.Replace("$DepartDate$", "<i>" & Format(dDepartureDate, "dddd, MMMM ") & dDepartureDate.Day & ", " & dDepartureDate.Year & "<i>")
      sTemplate = sTemplate.Replace("$RentalRate$", FormatCurrency(oBook.Rate__Numeric, 2))
      sTemplate = sTemplate.Replace("$TaxRate$", FormatPercent(oProp.TaxRate__Integer / 100, 0))
      sTemplate = sTemplate.Replace("$TaxDue$", FormatCurrency(oBook.TaxDue__Numeric, 2))
      sTemplate = sTemplate.Replace("$RentalTotal$", FormatCurrency(dTotalPaymentDue, 2))


      sTemplate = sTemplate.Replace("$FirstPaymentDueDate$", sFirstPaymentDueDate)
      sTemplate = sTemplate.Replace("$FirstPaymentDue$", sFirstPaymentDue)
      sTemplate = sTemplate.Replace("$FirstPaymentReceived$", FormatCurrency(dFirstPaymentReceived, 2))
      sTemplate = sTemplate.Replace("$FirstPaymentRemaining$", FormatCurrency(dFirstPaymentRemaining, 2))

      sTemplate = sTemplate.Replace("$BalancePaymentDueDate$", sBalancePaymentDueDate)
      sTemplate = sTemplate.Replace("$BalancePaymentDue$", sBalancePaymentDue)
      sTemplate = sTemplate.Replace("$BalancePaymentReceived$", FormatCurrency(dBalancePaymentReceived, 2))
      sTemplate = sTemplate.Replace("$BalancePaymentRemaining$", FormatCurrency(dBalancePaymentRemaining, 2))

      sTemplate = sTemplate.Replace("$TotalPaymentDue$", FormatCurrency(dTotalPaymentDue, 2))
      sTemplate = sTemplate.Replace("$TotalPaymentReceived$", FormatCurrency(dTotalPaymentReceived, 2))
      sTemplate = sTemplate.Replace("$TotalPaymentRemaining$", sTotalRemaining)


      If oProp.DamageDeposit__Numeric > 0 Then
        sTemplate = sTemplate.Replace("$DamageDeposit$", FormatCurrency(oProp.DamageDeposit__Numeric, 2))
      Else
        sTemplate = ClipHTMLSectionToUse(sTemplate, "<DAMAGE_DEPOSIT>", True)
        sTemplate = ClipHTMLSectionToUse(sTemplate, "<damage_deposit>", True)
      End If
      sResult = sTemplate & sTemp
      CreateReservationSummary = True

      oDB = Nothing
    Catch ex As Exception
      If bPutErrorInResult Then sResult = "ERROR:" & vbCrLf & "Message: " & ex.Message.ToString & vbCrLf & "Site:  " & ex.TargetSite.ToString & vbCrLf & "Stack Trace:  " & ex.StackTrace.ToString
      oErr.HandleError(ex, False)
    End Try

  End Function


  Public Function GetHostPaymentsBalance(ByRef oTrans As System.Data.SqlClient.SqlTransaction, ByVal lHostID As Long, ByVal sType As String, ByVal lPropertyGroup As Long) As Double

    Try

      GetHostPaymentsBalance = 0
      Dim oDB As New DBUtilities

      If lHostID > 0 Then
        If lPropertyGroup > 0 Then
          oDB.SelectData("Select Balance from vwHost" & sType & "PaymentsBalance where HostID=" & lHostID & " And IsNull(PropertyGroup,1)=" & lPropertyGroup, , , oTrans, True)
        Else
          oDB.SelectData("Select Balance from vwHost" & sType & "PaymentsBalance where HostID=" & lHostID, , , oTrans, True)
        End If
      Else
        If lPropertyGroup > 0 Then
          oDB.SelectData("Select sum(Balance) as Balance from vwHost" & sType & "PaymentsBalance Where IsNull(PropertyGroup,1)=" & lPropertyGroup, , , oTrans, True)
        Else
          oDB.SelectData("Select sum(Balance) as Balance from vwHost" & sType & "PaymentsBalance ", , , oTrans, True)
        End If
      End If

      If oDB.MoveFirst Then
        GetHostPaymentsBalance = oDB.CurrentRow("Balance")
      End If

    Catch ex As Exception
      oErr.HandleError(ex, False)
    End Try

  End Function



  Public Function ApplyPaymentCredit(ByRef oBook As TableBookings, ByRef oHostPay As TableHostPayments, ByVal lHostID As Long, ByVal lBookingID As Long, ByVal dTotalDepAmount As Double, dDepAmountToCreateCreditsFor As Double, ByVal dTotalBalAmount As Double, dBalAmountToCreateCreditsFor As Double, ByRef cDepAmountCredited As Double) As Boolean

    ' BJR IMPORTANT - If anything changes in this function, must change counterpart in Rental software


    ApplyPaymentCredit = False
    Try

      Dim sType As String = "", lPropertyGroup As Long, oDB As New DBUtilities, oDB2 As New DBUtilities, cBalance As Double, bKeepDepsoitBalanceSeparate As Boolean = False
      Dim sCreditAmount As String, cCreditAmount As Double, dDepositCreditsAlreadyUsed As Double, dBalanceCreditsAlreadyUsed As Double

      If lHostID > 0 And lBookingID > 0 Then


        lPropertyGroup = 1
        oDB.SelectData("Select PropertyGroup from vwBookingProperty where Booking_id=" & lBookingID, , , oBook.Transaction, True)
        If oDB.MoveFirst Then lPropertyGroup = CNullI(oDB.CurrentRow("PropertyGroup"))
        If lPropertyGroup = 0 Then lPropertyGroup = 1

        If dDepAmountToCreateCreditsFor > 0 Then

          If bKeepDepsoitBalanceSeparate Then sType = "Deposit"
          cBalance = GetHostPaymentsBalance(oBook.Transaction, lHostID, sType, lPropertyGroup)
          If cBalance > 0 Then

            ' Get total deposit credits already given to this booking.  Total credits can't exceed total dep required
            oDB.SelectData("Select Isnull(TotalCredits,0) as TotalCredits from vwHostPaymentCreditsUsedByBooking where PaymentType='First Payment' and  ToBookingID=" & lBookingID)
            If oDB.MoveFirst Then dDepositCreditsAlreadyUsed = oDB.CurrentRow("TotalCredits")

            If dTotalDepAmount - dDepositCreditsAlreadyUsed < dDepAmountToCreateCreditsFor Then
              dDepAmountToCreateCreditsFor = dTotalDepAmount - dDepositCreditsAlreadyUsed
            End If

            ' Get all Bookings with deposit credits yet to give
            oDB.SelectData("Select * from vwHostPayments" & sType & "BalanceByBooking where HostID=" & lHostID & " And PropertyGroup=" & lPropertyGroup, , , oBook.Transaction, True)

            If oDB.MoveFirst Then

              Dim iFromBookingID As Integer, cAmountRemaining As Double

              ' Apply to deposit first
              cAmountRemaining = dDepAmountToCreateCreditsFor

              ' Go through all the booking credits until run out of credits to apply, or this host payment has been paid
              Do
                cCreditAmount = oDB.CurrentRow("Balance")
                If cAmountRemaining <= cCreditAmount Then cCreditAmount = cAmountRemaining
                cAmountRemaining = cAmountRemaining - cCreditAmount
                iFromBookingID = oDB.CurrentRow("FromBookingID")
                oHostPay.Clear()
                oHostPay.HostID_RQ__Integer = lHostID
                oHostPay.FromBookingID__Integer = iFromBookingID
                oHostPay.ToBookingID__Integer = lBookingID
                oHostPay.Amount_RQ__Numeric = -1 * cCreditAmount
                oHostPay.Username__String = "Online"
                oHostPay.TransactionDate_RQ__Date = Now
                oHostPay.TransactionType_RQ__String = "Applied To First Payment Due Host"
                oHostPay.PropertyGroup__Integer = lPropertyGroup
                oHostPay.Insert()
                cDepAmountCredited = cDepAmountCredited + cCreditAmount
              Loop While cAmountRemaining > 0 And oDB.MoveNext

              '              ApplyPaymentCredit = cDepAmountCredited
            End If
          End If
        End If

        If dBalAmountToCreateCreditsFor > 0 Then

          If bKeepDepsoitBalanceSeparate Then sType = "Balance"
          cBalance = GetHostPaymentsBalance(oBook.Transaction, lHostID, sType, lPropertyGroup)
          If cBalance > 0 Then

            ' Get total balance credits already given to this booking.  Total credits can't exceed total bal required
            oDB.SelectData("Select Isnull(TotalCredits,0) as TotalCredits from vwHostPaymentCreditsUsedByBooking where PaymentType='Balance Payment' and  ToBookingID=" & lBookingID)
            If oDB.MoveFirst Then dBalanceCreditsAlreadyUsed = oDB.CurrentRow("TotalCredits")

            If dTotalBalAmount - dBalanceCreditsAlreadyUsed < dBalAmountToCreateCreditsFor Then
              dBalAmountToCreateCreditsFor = dTotalBalAmount - dBalanceCreditsAlreadyUsed
            End If


            ' Get all Bookings with balance credits yet to give
            oDB.SelectData("Select * from vwHostPayments" & sType & "BalanceByBooking where HostID=" & lHostID & " And PropertyGroup=" & lPropertyGroup, , , oBook.Transaction, True)

            If oDB.MoveFirst Then

              Dim iFromBookingID As Integer, cAmountRemaining As Double

              ' Apply to balance 
              cAmountRemaining = dBalAmountToCreateCreditsFor

              ' Go through all the booking credits until run out of credits to apply, or this host payment has been paid
              Do
                cCreditAmount = oDB.CurrentRow("Balance")
                If cAmountRemaining <= cCreditAmount Then cCreditAmount = cAmountRemaining
                cAmountRemaining = cAmountRemaining - cCreditAmount
                iFromBookingID = oDB.CurrentRow("FromBookingID")
                oHostPay.Clear()
                oHostPay.HostID_RQ__Integer = lHostID
                oHostPay.FromBookingID__Integer = iFromBookingID
                oHostPay.ToBookingID__Integer = lBookingID
                oHostPay.Amount_RQ__Numeric = -1 * cCreditAmount
                oHostPay.Username__String = "Online"
                oHostPay.TransactionDate_RQ__Date = Now
                oHostPay.TransactionType_RQ__String = "Applied To First Payment Due Host"
                oHostPay.PropertyGroup__Integer = lPropertyGroup
                oHostPay.Insert()
                cDepAmountCredited = cDepAmountCredited + cCreditAmount
              Loop While cAmountRemaining > 0 And oDB.MoveNext

            End If
          End If
        End If

        ApplyPaymentCredit = True
      Else
        oErr.LogWebError("Host or Booking ID not given, HostID=" & lHostID & ", BookingID=" & lBookingID, "SHaredUtilities:ApplyPaymentCredit")
      End If

    Catch ex As Exception
      oErr.HandleError(ex, False)
      Err.Clear()
    End Try

  End Function


  Public Function SavePayment(ByRef sCSTAYResult As String, ByRef sUserResult As String, ByRef oBook As TableBookings, ByRef oPayment As Object, ByRef oGuest As TableGuests,
  ByRef oProp As TableProperties, ByRef oHost As TableHosts, ByRef oHostPay As TableHostPayments, ByRef oCIMGate As TableCIMGatewayActivity,
  ByRef sCustomerCIMID As String, ByRef sPaymentCIMID As String, ByVal iBookingID As Integer,
  ByRef iPaymentID As Integer, ByVal sPaymentOrigin As String, ByVal sPaymentLocation As String, ByVal sEmailType As String, ByVal bSendEmails As Boolean, ByVal bChargeInTestMode As Boolean,
  ByVal dPaymentAmount As Double, ByVal dChargeAmount As Double, ByVal sPaymentCategory As String, ByVal sPaymentMethod As String,
  ByVal sPaymentEmail As String, ByRef iContactLogID As Integer, ByVal bFinalBalPayMade As Boolean, ByVal bFinalDepPayMade As Boolean,
  ByRef oAuth As CStayAuthorizeNET, ByVal sEncryptPassword As String, ByVal sCCType As String, ByVal sCCNumber As String, ByVal sCCFirstName As String, ByVal sCCLastName As String,
  ByVal sCCExpMonth As String, ByVal sCCExpYear As String, ByVal sCCVerification As String, ByVal sCCAddress As String, ByVal sCCCity As String,
  ByVal sCCState As String, ByVal sCCZip As String, ByVal sBankName As String, ByVal sBankAccountName As String, ByVal sBankAccountNumber As String,
  ByVal sBankABANumber As String, ByVal sBankAccountType As String, ByVal bSkipMakingCharge As Boolean, Optional ByVal bPutErrorInResult As Boolean = False,
  Optional ByVal sCreditType As String = "", Optional ByVal sPrevTransactionID As String = "", Optional ByVal iPrevPaymentID As Integer = 0, Optional ByVal sPrevPaymentCategory As String = "", Optional ByRef oPrevPayment As Object = Nothing,
  Optional ByRef sAuthorizationTypeUsed As String = "", Optional ByVal sUserName As String = "", Optional ByRef bProblemCard As Boolean = False, Optional ByVal bUseCIMGateway As Boolean = True, Optional ByVal sConnString As String = "", Optional ByRef bCIMTransactionDeclined As Boolean = False, Optional ByVal bQueueEmail As Boolean = False) As Boolean


    SavePayment = False

    Try

      If sPaymentCategory.ToString = "" Then
        Throw New Exception("Payment type was not specified")
        Exit Function
      End If

      Dim oAES As New MyAES
      Dim dCurrBalAmount As Double = 0, sThisCSTAYResult As String = "", sEmailToName As String = "", sCurrPaymentNotes As String = "", dCurrPaymAmount As Double = 0
      Dim dCurrDepAmount As Double = 0, sPropImgURL As String = "", dHostPaidAmtDep As Double = 0, dHostPaidAmtBal As Double = 0, sResult As String = "", iLateCancelBookID As Integer = 0
      Dim sConfirmEmail As String = "", sAdminConfirmEmail As String = "", dHostPymtAmt As Double = 0, bRecordPayment As Boolean = True, sPostURL As String = ""
      Dim sGuestEmail2 As String = "", sHostEmail2 As String = "", sHostEmail As String = "", sHostName As String = "", iHostID As Integer = 0, cDepAmountCredited As Double = 0

      If iBookingID = 0 Then
        Throw New Exception("Reservation ID was not specified")
        Exit Function
      Else

        oBook.Booking_ID_PK__Integer = iBookingID
        oBook.SelectData()
        If oBook.Booking_ID_PK__Integer = 0 Then
          Throw New Exception("Reservation record not available for ReservationID=" & iBookingID)
          Exit Function
        Else
          iHostID = oBook.Host_ID__Integer

          If sPaymentCategory.ToString = "Balance Due" Then
            oPayment = New TableBalancePayments(oBook.Transaction)
          ElseIf sPaymentCategory.ToString = "Deposit" Then
            oPayment = New TableDepositPayments(oBook.Transaction)
          Else
            oPayment = New TableRefundPayments(oBook.Transaction)
          End If

          dHostPaidAmtDep = oBook.HostPaidAmtDep__Numeric
          If dHostPaidAmtDep = 0 Then dHostPaidAmtDep = oBook.Rate__Numeric - oBook.Commission__Numeric
          dHostPaidAmtBal = oBook.HostPaidAmtBal__Numeric
          If dHostPaidAmtBal = 0 Then dHostPaidAmtBal = oBook.BalanceDue__Numeric

          ' Create payment record 
          oPayment.BookingID__Integer = iBookingID
          oPayment.GuestID__Integer = oBook.Guest_ID__Integer
          oPayment.PropertyID__Integer = oBook.Property_ID__Integer
          oPayment.HostID__Integer = iHostID
          oPayment.Category__String = sPaymentCategory

          oPayment.Location__String = sPaymentLocation
          oPayment.Amount__Numeric = dPaymentAmount
          oPayment.Date__Date = Now

          If sPaymentMethod.ToString = "Card" Then
            oPayment.Type__String = sPaymentMethod
            oPayment.CCType__String = sCCType
            oPayment.CCName__String = sCCFirstName & " " & sCCLastName
            If sCCNumber.Length > 4 Then
              oPayment.CCNumber__String = sCCNumber.Substring(sCCNumber.Length - 4, 4)
              '              oPayment.CCNumberEncrypted__String = oAES.AES_Encrypt(sCCNumber, "freed0m")
            End If
            oPayment.CCExpMonth__String = sCCExpMonth
            oPayment.CCExpYear__String = sCCExpYear
            oPayment.CCAddress__String = sCCAddress
            oPayment.CCCity__String = sCCCity
            oPayment.CCState__String = sCCState
            oPayment.CCZip__String = sCCZip
            sEmailToName = sCCFirstName & " " & sCCLastName
          Else
            oPayment.Type__String = "eCheck"
            oPayment.BankAccountName__String = sBankAccountName
            oPayment.BankAccountNumber__String = sBankAccountNumber.Substring(sBankAccountNumber.Length - 4, 4)
            oPayment.BankAccountType__String = sBankAccountType
            oPayment.BankName__String = sBankName
            oPayment.BankRoutingNumber__String = sBankABANumber
            sEmailToName = sBankAccountName
          End If

          oPayment.Email__String = sPaymentEmail

          oPayment.Status__String = "Charged"
          If sPaymentCategory = "Balance Due" Then oPayment.UseOnReceipt__Integer = 1
          oPayment.CCVerification__String = sCCVerification
          iPaymentID = oPayment.Insert


          If iPaymentID = 0 Then
            Throw New Exception("Could not add Payment record")
          Else

            oGuest.Guest_ID_PK__Integer = oBook.Guest_ID__Integer
            oGuest.SelectData()
            If sCustomerCIMID = "" Then sCustomerCIMID = oGuest.AuthNETCustomerProfileID__String
            sGuestEmail2 = oGuest.Email2__String
            oProp.Property_ID_PK__Integer = oBook.Property_ID__Integer
            oProp.SelectData()
            oHost.Host_ID_PK__Integer = iHostID
            oHost.SelectData()
            ' Only send email to host if reservation made, not for payments
            If sPaymentOrigin = "Reservation" Then
              sHostName = oHost.HostName__String
              sHostEmail = oHost.Email__String
              sHostEmail2 = oHost.Email2__String
            End If
            If bSendEmails Then
              If Not CreateEmailHTML(sCSTAYResult, sPaymentOrigin, sPaymentCategory, sEmailType, sConfirmEmail, sAdminConfirmEmail, iBookingID, "", oPayment, oBook, oGuest, oProp, oHost) Then oErr.LogWebError("Unable to Send Email: " & sCSTAYResult, "SharedUtilities:SavePayment")
            End If

            If oAuth.ProcessCharge(oAuth, oBook, oPayment, oGuest, oProp, oCIMGate, oHost, dChargeAmount, sPaymentOrigin, sPaymentMethod, sPaymentLocation,
            sCustomerCIMID, sPaymentCIMID, sCCNumber, sCCVerification,
            CNullI(sCCExpMonth), CNullI(sCCExpYear), CNullS(sCCFirstName),
            CNullS(sCCLastName), CNullS(sCCAddress), CNullS(sCCCity),
            CNullS(sCCState), CNullS(sCCZip), sBankABANumber, sBankAccountName,
            sBankAccountNumber, sBankAccountType, sBankName, sCSTAYResult, sUserResult, bChargeInTestMode, bPutErrorInResult,
            bSkipMakingCharge, sCreditType, sPrevTransactionID, sPostURL, bUseCIMGateway, bCIMTransactionDeclined) Then

              ProcessTransactionAndReset(True, sPaymentCategory, sPrevPaymentCategory, oBook, oPayment, oPrevPayment, oGuest, oProp, oHost, oHostPay, oCIMGate, sConnString)

              Try

                sConfirmEmail = sConfirmEmail.Replace("$ChargeUnderReviewMessage$", oAuth.UserChargeResultText)
                sAdminConfirmEmail = sAdminConfirmEmail.Replace("$ChargeUnderReviewMessage$", oAuth.UserChargeResultText)
                oPayment.Clear()
                oPayment.PaymentID_PK__Integer = iPaymentID
                sAuthorizationTypeUsed = oAuth.InpAuthorizationType.ToString


                If sCreditType <> "" Then

                  If oAuth.InpAuthorizationType.ToString <> "VOID" And sCreditType <> "Void" Then
                    oPayment.Status__String = "Credited"
                    oPayment.OriginalPaymentCategory__String = sPrevPaymentCategory
                    oPayment.OriginalPaymentID__Integer = iPrevPaymentID
                  End If

                  ' If a prev credit card charge was voided, mark the prev payment records as voided
                  oPrevPayment.PaymentID_PK__Integer = iPrevPaymentID
                  oPrevPayment.SelectData()
                  If oPrevPayment.MoveFirst Then
                    sCurrPaymentNotes = oPrevPayment.Notes__String.ToString
                    dCurrPaymAmount = oPrevPayment.Amount__Numeric
                    oPrevPayment.Clear()
                    oPrevPayment.PaymentID_PK__Integer = iPrevPaymentID
                    If sCreditType = "Void" Or oAuth.InpAuthorizationType.ToString = "VOID" Then
                      oPrevPayment.Status__String = "Voided"
                      oPayment.Status__String = "VoidRefund"
                    Else
                      oPrevPayment.Status__String = "Refunded"
                    End If
                    oPrevPayment.Notes__String = sCurrPaymentNotes & "Payment Of " & FormatCurrency(dCurrPaymAmount.ToString) & " " & oPrevPayment.Status__String & " by " & sUserName & " On " & Now & vbCrLf
                    oPrevPayment.Update()
                    oPayment.Update()
                  Else
                    oErr.LogWebError("Could Not Get prev payment record For " & sCreditType & " credit card transaction " & vbCrLf & "Credit Card TransID: " & sPrevTransactionID & vbCrLf & "PaymentID: " & iPrevPaymentID & vbCrLf & "Username: " & sUserName, "SharedUtilities:SavePayment")
                  End If


                End If


                If bChargeInTestMode Then oAuth.RespApprovalCode = "TEST MODE"
                oPayment.AuthorizeNetApprovalCode__String = oAuth.RespApprovalCode
                oPayment.AuthorizeNetAVSResultCode__String = oAuth.RespAVSCode
                oPayment.AuthorizeNetAVSResultText__String = oAuth.RespAVSText
                oPayment.AuthorizeNetCVCCResponseCode__String = oAuth.RespCVCCCode
                oPayment.AuthorizeNetCVCCResponseText__String = oAuth.RespCVCCText
                oPayment.AuthorizeNetReasonCode__String = oAuth.RespReasonCode
                oPayment.AuthorizeNetReasonText__String = oAuth.RespReasonText
                oPayment.AuthorizeNetReturnCode__String = oAuth.RespCode
                oPayment.AuthorizeNetTransactionID__String = oAuth.RespTransID
                oPayment.AuthorizeNetURL__String = sPostURL
                oPayment.Update()

                If sCreditType = "" Then

                  ' When issuing credit, don't need to update Booking record
                  oPayment.Clear()
                  oPayment.PaymentID_PK__Integer = iPaymentID
                  oPayment.SelectData()

                  oBook.Clear()
                  oBook.Booking_ID_PK__Integer = iBookingID
                  oBook.SelectData()
                  bFinalBalPayMade = ((sPaymentCategory = "Balance Due") And (oBook.BalanceDue__Numeric - dPaymentAmount) = 0.0) Or (oBook.BalanceDue__Numeric = 0)
                  oBook.Clear()
                  oBook.Booking_ID_PK__Integer = iBookingID

                  If sPaymentCategory = "Balance Due" Then
                    oBook.BalanceAmount__Numeric = dCurrBalAmount + dPaymentAmount
                    oBook.BalCardType__String = sCCType
                    oBook.BalCardNumber__String = Right(sCCNumber, 4)
                    If bFinalBalPayMade Then
                      oBook.FinalBalancePaymentMadeOnline__Integer = bFinalBalPayMade
                      ' Status is changed manually if order taken by phone
                      If sPaymentLocation <> "Phone" Then oBook.Status__String = "Balance Received"

                      oBook.BalanceReceived__Date = FormatDateTime(Now, DateFormat.ShortDate)
                      oBook.HostPdDate__Date = Now
                      oBook.HostPaidAmtBal__Numeric = oBook.BalanceDue__Numeric
                      oBook.BalCardConfirm__String = oAuth.RespApprovalCode
                      ' Don't apply credit if payment being made from Rental SW i.e. sUserName is supplied
                      If dHostPaidAmtBal > 0 And sUserName = "" Then
                        '                        If dPaymentAmount = oBook.BalanceDue__Numeric Or dPaymentAmount = oBook.BalanceAmount__Numeric Then
                        If Not ApplyPaymentCredit(oBook, oHostPay, iHostID, iBookingID, 0, 0, dHostPaidAmtBal, dHostPaidAmtBal, cDepAmountCredited) Then oErr.LogWebError("Unable To ApplyPaymentCredit For Balance Payment online For BookingID=" & iBookingID)
                      End If
                    Else
                      ' Status is changed manually if order taken by phone
                      If sPaymentLocation <> "Phone" Then oBook.Status__String = "First Payment Received"
                    End If
                  Else
                    oBook.DepositAmountPaid__Numeric = dCurrDepAmount + dPaymentAmount
                    oBook.DepCardConfirm__String = oAuth.RespApprovalCode
                    oBook.DepositReceived__Date = Now.ToShortDateString
                    oBook.DepCardType__String = sCCType
                    oBook.DepCardNumber__String = Right(sCCNumber, 4)
                    If bFinalDepPayMade Or bFinalBalPayMade Then
                      oBook.FinalDepositPaymentMadeOnline__Integer = bFinalDepPayMade
                      oBook.DepositReceived__Date = FormatDateTime(GetDate, DateFormat.ShortDate)
                      ' Don't apply credit if payment being made from Rental SW i.e. sUserName is supplied
                      If dHostPaidAmtDep > 0 And sUserName = "" Then
                        '                        If dPaymentAmount = oBook.DepositAmount__Numeric Or dPaymentAmount = oBook.DepositAmountPaid__Numeric Then
                        If Not ApplyPaymentCredit(oBook, oHostPay, iHostID, iBookingID, dHostPaidAmtDep, dHostPaidAmtDep, 0, 0, cDepAmountCredited) Then oErr.LogWebError("Unable To ApplyPaymentCredit For First Payment online For BookingID=" & iBookingID)
                      End If

                      ' if making full payment up front
                      If bFinalBalPayMade Then
                        ' Status is changed manually if order taken by phone
                        If sPaymentLocation <> "Phone" Then oBook.Status__String = "Balance Received"
                        oBook.BalanceReceived__Date = oBook.DepositReceived__Date
                      Else
                        If sPaymentLocation <> "Phone" Then oBook.Status__String = "First Payment Received"
                      End If
                    Else
                      If sPaymentLocation <> "Phone" Then oBook.Status__String = "Booked, Waiting for First Payment"
                    End If
                  End If

                  oBook.Update()
                  oBook.Booking_ID_PK__Integer = iBookingID
                  oBook.SelectData()

                  ' See if late cancellation made for reserved property.  If so, link it to this booking, create hostpayment credit, and send email to Lonetta
                  If sPaymentOrigin = "Reservation" Then
                    If PropertyAvailable(oBook.Property_ID__Integer, CDate(oBook.ArriveDate__Date), CDate(oBook.DepartDate__Date), sResult, True, iLateCancelBookID, iBookingID) And iLateCancelBookID > 0 Then
                      sAdminConfirmEmail = "<span style=""color: Red; font-weight:bold; font-size:18px"">LINKED TO LATE CANCELLATION RESERV# " & iLateCancelBookID & "</span>" & sAdminConfirmEmail
                    End If
                  End If

                  If bSendEmails Then
                    If sPaymentOrigin <> "Reservation" Then
                      sHostEmail = ""
                      sHostEmail2 = ""
                    End If

                    SendConfirmationEmail(sEmailType.Replace("Email", ""), sConfirmEmail, sAdminConfirmEmail, sPaymentEmail, sGuestEmail2, sEmailToName, iBookingID, iContactLogID, sHostName, sHostEmail, sHostEmail2, bQueueEmail)
                  End If
                End If

                SavePayment = True
              Catch ex As Exception
                If bPutErrorInResult Then sCSTAYResult &= sPaymentCategory & "  " & oPayment.ConnectionString & " : " & oGuest.ConnectionString & "   ERROR:" & vbCrLf & "Message: " & ex.Message.ToString & vbCrLf & "Site:  " & ex.TargetSite.ToString & vbCrLf & "Stack Trace:  " & ex.StackTrace.ToString
                oErr.HandleError(ex, False)
              End Try

            Else

              If oAuth.RespCode = "2" Or bCIMTransactionDeclined Then
                bProblemCard = True
                Try
                  oAuth.SendProblemCardEmail(oPayment, oAuth, oHost, oGuest, oProp)
                Catch ex As Exception
                  oErr.HandleError(ex, "SharedUtilitiesSavePayment:SendProblemCardEmail", False)
                End Try
              End If
              ProcessTransactionAndReset(False, sPaymentCategory, sPrevPaymentCategory, oBook, oPayment, oPrevPayment, oGuest, oProp, oHost, oHostPay, oCIMGate, sConnString)

            End If
            sCSTAYResult &= oAuth.CStayChargeResultText
            sUserResult = oAuth.UserChargeResultText
          End If

        End If
      End If
    Catch ex As Exception
      oErr.HandleError(ex, False)
      If bPutErrorInResult Then
        sCSTAYResult &= "ERROR: " & vbCrLf & "Message: " & ex.Message.ToString & vbCrLf & "Site:  " & ex.TargetSite.ToString & vbCrLf & "Stack Trace:  " & ex.StackTrace.ToString
      Else
        sCSTAYResult &= "A technical problem ocurred,<BR/>And we could Not complete your " & sPaymentOrigin & " at this time.<BR/>Please try again later Or call <span class=""SmallTitle"">607-547-6260</span> .<BR/>"
      End If
    End Try


  End Function

  Private Sub ProcessTransactionAndReset(ByVal bSaveTransaction As Boolean, ByVal sPaymentCategory As String, ByVal sPrevPaymentCategory As String, ByRef oBook As TableBookings, ByRef oPayment As Object, ByRef oPrevPayment As Object, ByRef oGuest As TableGuests, _
  ByRef oProp As TableProperties, ByRef oHost As TableHosts, ByRef oHostPay As TableHostPayments, ByRef oCIMGate As TableCIMGatewayActivity, ByVal sConnString As String)

    Try

      oBook.ProcessTransaction(bSaveTransaction)

      oBook = Nothing
      oGuest = Nothing
      oHost = Nothing
      oHostPay = Nothing
      oProp = Nothing
      oPayment = Nothing
      oPrevPayment = Nothing
      oCIMGate = Nothing
      If sConnString = "" Then
        oBook = New TableBookings
        oGuest = New TableGuests
        oHost = New TableHosts
        oHostPay = New TableHostPayments
        oProp = New TableProperties
        oCIMGate = New TableCIMGatewayActivity

        If sPaymentCategory.ToString = "Balance Due" Then
          oPayment = New TableBalancePayments
        ElseIf sPaymentCategory.ToString = "Deposit" Then
          oPayment = New TableDepositPayments
        Else
          oPayment = New TableRefundPayments
        End If

        If sPrevPaymentCategory = "Deposit" Then
          oPrevPayment = New TableDepositPayments
        ElseIf sPrevPaymentCategory = "Balance Due" Then
          oPrevPayment = New TableBalancePayments
        End If
      Else
        oBook = New TableBookings(sConnString)
        oGuest = New TableGuests(sConnString)
        oHost = New TableHosts(sConnString)
        oHostPay = New TableHostPayments(sConnString)
        oProp = New TableProperties(sConnString)
        oCIMGate = New TableCIMGatewayActivity(sConnString)
        If sPaymentCategory.ToString = "Balance Due" Then
          oPayment = New TableBalancePayments(sConnString)
        ElseIf sPaymentCategory.ToString = "Deposit" Then
          oPayment = New TableDepositPayments(sConnString)
        Else
          oPayment = New TableRefundPayments(sConnString)
        End If

        If sPrevPaymentCategory = "Deposit" Then
          oPrevPayment = New TableDepositPayments(sConnString)
        ElseIf sPrevPaymentCategory = "Balance Due" Then
          oPrevPayment = New TableBalancePayments(sConnString)
        End If
      End If
    Catch ex As Exception
      oErr.HandleError(ex, False)
    End Try

  End Sub


  Public Sub SendConfirmationEmail(ByVal sType As String, ByRef sEmail As String, ByRef sAdminEmail As String, ByVal sGuestEmail As String, ByVal sGuestEmail2 As String, _
  ByVal sGuestEmailName As String, ByVal iBookingID As Integer, ByRef iContactLogID As Long, ByVal sHostEmailName As String, ByVal sHostEmail As String, ByVal sHostEmail2 As String, ByVal bQueueEmail As Boolean)

    Try

      If sGuestEmail.ToString <> "" Then
        If sGuestEmailName = "" Then sGuestEmailName = sGuestEmail
        Dim oDB As New DBUtilities, sErrorWhileSendMsg As String = ""
        If Not SendMail(sGuestEmail.ToString, sSystemEmailAddress, "Cooperstown Stay - Online " & sType, sEmail, "Cooperstown Stay", sGuestEmailName.ToString, True, bQueueEmail, , , "Bookings", iBookingID) Then
          oErr.LogWebError("Could Not send the following " & sType & " confirmation email to " & sGuestEmail & vbCrLf & vbCrLf & sEmail, "Reserve.aspx:CreateReservationConfirmation")
          sErrorWhileSendMsg = "ERR SEND EMAIL FROM WEBSITE-"
        End If
        ' Need to add contact even if error sending email so don;t get error on reservecomplete.aspx page
        oDB.SelectData("EXEC spAddBookingContact " & iBookingID & ", Null, " & FixParam(sGuestEmail.ToString, True) & ", " & FixParam(sType & " Email To Guest", True) & ", Null, " & FixParam(sEmail, True) & ",'Cooperstown Stay - Online " & sType & "','html'", sConnString)
        If oDB.MoveNext Then iContactLogID = oDB.CurrentRow("LogID")
        If sGuestEmail2.Trim <> "" Then SendMail(sGuestEmail2.ToString, sSystemEmailAddress, sErrorWhileSendMsg & "Cooperstown Stay - Online " & sType, sEmail, "Cooperstown Stay", sGuestEmailName.ToString, True, bQueueEmail, , , "Bookings", iBookingID)

                      If sHostEmail.Trim <> "" Then
                        If Not SendMail(sHostEmail.ToString, sSystemEmailAddress, "Cooperstown Stay " & sType, sEmail, "Cooperstown Stay", sHostEmailName.ToString, True, bQueueEmail, , , "Bookings", iBookingID) Then
                          oErr.LogWebError("Could not send the following " & sType & " confirmation email to :" & sHostEmail & vbCrLf & vbCrLf & sEmail, "Reserve.aspx:CreateReservationConfirmation")
                        Else
                          oDB.SelectData("EXEC spAddBookingContact " & iBookingID & ",Null," & FixParam(sHostEmail.ToString, True) & "," & FixParam(sType & " Email To Host", True) & ",Null," & FixParam(sEmail, True) & ",'Cooperstown Stay - Online " & sType & "','html'", sConnString)
                        End If
                        If sHostEmail2.Trim <> "" Then SendMail(sHostEmail2.ToString, sSystemEmailAddress, "Cooperstown Stay " & sType, sEmail, "Cooperstown Stay", sHostEmailName.ToString, True, bQueueEmail, , , "Bookings", iBookingID)
                      End If

                      If Not SendMail(sSystemEmailAddress, sSystemEmailAddress, "Online " & sType & " - " & iBookingID, sAdminEmail, "Cooperstown Stay", sSystemEmailAddress, True, bQueueEmail, , , "Bookings", iBookingID) Then
                        oErr.LogWebError("Could not send the following " & sType & " confirmation email to " & sSystemEmailAddress & ": " & vbCrLf & vbCrLf & sAdminEmail, "Reserve.aspx:CreateReservationConfirmation")
                      End If
                      oDB = Nothing
                    End If
    Catch ex As Exception
      oErr.HandleError(ex, False, False)
    End Try
  End Sub







#End Region

#Region "Shared Utilities"


  Public Function GetDate() As Date
    GetDate = Now
    If IsDate(sDateForTesting) Then GetDate = CDate(sDateForTesting)
  End Function

  Public Shared Function GetDateS() As Date
    GetDateS = Now
    If IsDate(sDateForTesting) Then GetDateS = CDate(sDateForTesting)
  End Function

  Public Function GetCodeValue(ByVal sCodeKey As String, ByVal bUseLargeValue As Boolean, Optional ByVal oCode As TableCodeData = Nothing, Optional ByVal bClipHTMLSection As Boolean = True) As String

    GetCodeValue = ""
    Try

      If oCode Is Nothing Then oCode = New TableCodeData(sConnString)
      oCode.Clear()
      oCode.CodeKey__String = sCodeKey
      oCode.SelectData()
      If oCode.CodeID_PK__Integer > 0 Then
        If bUseLargeValue Then
          GetCodeValue = oCode.CodeValueLarge__String
        Else
          GetCodeValue = oCode.CodeValue__String
        End If
      End If
      If bClipHTMLSection Then GetCodeValue = ClipHTMLSectionToUse(GetCodeValue, "$$EditThisSectionOnly$$")
    Catch ex As Exception
      oErr.HandleError(ex, False)
    End Try

  End Function

  Public Function ClipHTMLSectionToUse(ByVal sInput As String, ByVal sClipKey As String, Optional ByVal bEliminateClippedSection As Boolean = False, Optional ByVal sStringToSubstitue As String = "") As String

    ClipHTMLSectionToUse = sInput
    Dim sData() As String, iIndex As Integer = 0, sOut As String = "", iStart As Integer = 1, iStop As Integer = 0, iStep As Integer = 2
    Try

      sData = Split(sInput, sClipKey)
      If UBound(sData) > 1 Then
        iStop = UBound(sData)
        If bEliminateClippedSection Then
          iStart = 0
          If sStringToSubstitue <> "" Then
            iStep = 1
          Else
            If iStop > 2 Then iStop = iStop - 2
          End If
        End If
        For iIndex = iStart To iStop Step iStep
          If iStep = 1 And (((iIndex + 2) Mod 2) <> 0) Then
            sOut &= sStringToSubstitue
          Else
            sOut &= ClipHTMLSectionToUse(sData(iIndex), sClipKey, bEliminateClippedSection, sStringToSubstitue)
          End If
        Next
        ClipHTMLSectionToUse = sOut
      End If
    Catch ex As Exception
      oErr.HandleError(ex, False)
    End Try

  End Function


  Public Function SendMailVB6(ByRef sResult As String, _
  ByVal sInSMTPServer As String, _
  ByVal sInSMTPUser As String, _
  ByVal sInSMTPPassword As String, _
  ByVal sFrom As String, _
  ByVal sFromName As String, _
  ByVal sTo As String, _
  ByVal sToName As String, _
  ByVal sSubject As String, _
  ByVal sBody As String, _
  ByVal BodyFormatIsHTML As Boolean, _
  ByVal bQueueEmail As Boolean, _
 ByVal sSourceTable As String, _
 ByVal iSourceID As Integer,
 ByVal iSMTPPort As Integer) As Boolean


    sSMTPServer = sInSMTPServer
    sSMTPUser = sInSMTPUser
    sSMTPPass = sInSMTPPassword
    SendMailVB6 = SendMail(sTo, sFrom, sSubject, sBody, sFromName, sToName, BodyFormatIsHTML, bQueueEmail, , sResult, sSourceTable, iSourceID, iSMTPPort)

  End Function



  Public Function SendMail( _
  ByVal sTo As String, _
  ByVal sFrom As String, _
  ByVal sSubject As String, _
  ByVal sBody As String, _
  Optional ByVal sFromName As String = "", _
  Optional ByVal sToName As String = "", _
  Optional ByVal BodyFormatIsHTML As Boolean = False, _
  Optional ByVal bQueueEmail As Boolean = False, _
  Optional ByVal bFromErrorHandler As Boolean = False, Optional ByRef sResult As String = "", _
  Optional ByVal sSourceTable As String = "", Optional ByVal iSourceID As Integer = 0, Optional ByVal bResendOnFailure As Boolean = True, Optional iSMTPPort As Integer = 0) As Boolean

    SendMail = False

    Try

      Dim sStatus As String = ""
      Dim sSubj As String, sErrMsg As String = ""

      If sToName = "" Then sToName = sTo
      If sFromName = "" Then sFromName = sFrom

      sSubj = sSubject

      If sSendAllEmailAddress <> "" Or bLiveTesting Then
                bQueueEmail = False
        If sSendAllEmailAddress = "" Then
        sTo = sErrorEmailAddress
        Else
          sTo = sSendAllEmailAddress
        End If
      End If
      If bQueueEmail Then sStatus = "Unsent-Queued"

      If (Not bQueueEmail) Then

        If sSMTPServer.ToString <> "" Then

          Dim oToAddr As New System.Net.Mail.MailAddress(sTo, sToName)
          Dim oFromAddr As New System.Net.Mail.MailAddress(sFrom, sFromName)
          Dim oMsg As New System.Net.Mail.MailMessage(oFromAddr, oToAddr)
          oMsg.Subject = sSubj
          oMsg.Body = sBody
          Dim oEmail As System.Net.Mail.SmtpClient

          If sSMTPServer.ToLower Like "*gmail*" Or sSMTPServer.ToLower Like "*amazonaws*" Then
            oEmail = New System.Net.Mail.SmtpClient(sSMTPServer, 587)
            oEmail.EnableSsl = True
          Else
            oEmail = New System.Net.Mail.SmtpClient(sSMTPServer)
          End If

          If sSMTPUser.ToString <> "" Then
            Dim SMTPUserInfo As New NetworkCredential(sSMTPUser, sSMTPPass)
            oEmail.DeliveryMethod = Mail.SmtpDeliveryMethod.Network
            oEmail.UseDefaultCredentials = False
            oEmail.Credentials = SMTPUserInfo
          End If
          If iSMTPPort > 0 Then
            oEmail.Port = iSMTPPort
            oEmail.EnableSsl = True
          End If


          oMsg.IsBodyHtml = BodyFormatIsHTML
          Try
            oEmail.Send(oMsg)
            SendMail = True
            sStatus = "Sent"
          Catch ex As Exception
            sStatus = "Unsent-Error While Sending"
            sResult &= "Error: " & ex.Message
            If ex.InnerException IsNot Nothing Then sResult &= "Inner: " & ex.InnerException.Message
            If Not bFromErrorHandler Then oErr.HandleError(ex, False, False)
            sErrMsg = sResult
          End Try
          oMsg = Nothing
          oEmail = Nothing
        End If
      End If


      Dim oMail As New TableQueuedEmail(sConnString)
      Try

        If Not bQueueEmail And Not bResendOnFailure And Not SendMail Then oMail.Tries__Integer = 100

        oMail.AddDate__Date = Now
        oMail.Body_RQ__String = sBody
        oMail.FormatIsHTML_RQ__Integer = BodyFormatIsHTML
        oMail.FromAddress_RQ__String = sFrom
        oMail.FromName__String = sFromName
        oMail.Priority__Integer = 5
        oMail.Status_RQ__String = sStatus
        '        oMail.Status_RQ__String = sStatus
        oMail.Subject_RQ__String = sSubj
        oMail.ToAddress_RQ__String = sTo
        oMail.ToName__String = sToName
        If sErrMsg <> "" Then oMail.ErrMessage__String = sErrMsg
        If sSourceTable <> "" Then oMail.LinkedTable__String = sSourceTable
        If iSourceID > 0 Then oMail.LinkedTableID__Integer = iSourceID
        oMail.SourceDB__String = oMail.ConnectionString
        If sStatus = "Sent" Then oMail.SendDate__Date = oMail.AddDate__Date

        If (oMail.Insert = 0) And bQueueEmail Then
          SendMail = False
        Else
          SendMail = True And (oMail.Tries__Integer = 0)
        End If

      Catch ex As Exception
        sResult &= "Error: " & ex.Message & "<br>"
        If ex.InnerException IsNot Nothing Then sResult &= "Inner: " & ex.InnerException.Message & "<br>"
        sResult &= "From: " & sFrom & "<br>To: " & sTo & "<br>Subject: " & sSubj
        If Not bFromErrorHandler Then oErr.LogWebError(sResult, "SendMail", , ex.TargetSite.ToString, , , ex.StackTrace)
      End Try

      oMail = Nothing
    Catch ex As Exception
      sResult = "Error: " & ex.Message
      If ex.InnerException IsNot Nothing Then sResult &= "Inner: " & ex.InnerException.Message
      sResult &= "From: " & sFrom & "<br>To: " & sTo & "<br>Subject: " & sSubject
      If Not bFromErrorHandler Then oErr.LogWebError(sResult, "SendMail", , ex.TargetSite.ToString, , , ex.StackTrace)

    End Try

  End Function




  Public Function CNullS(ByVal oInp As Object, Optional ByVal sDefault As String = "") As String
    CNullS = sDefault
    If Not IsDBNull(oInp) Then
      If Not (oInp Is Nothing) Then CNullS = CStr(oInp)
    End If
  End Function

  Public Function CNullI(ByVal oInp As Object, Optional ByVal iDefault As Integer = 0) As Integer
    CNullI = iDefault
    If Not IsDBNull(oInp) Then
      If Not (oInp Is Nothing) Then
        If IsNumeric(oInp) Then CNullI = CInt(oInp)
      End If
    End If
  End Function

  Public Function CNullD(ByVal oInp As Object, Optional ByVal dDefault As Double = 0) As Double
    CNullD = dDefault
    If Not IsDBNull(oInp) Then
      If Not (oInp Is Nothing) Then
        If IsNumeric(oInp) Then CNullD = CDbl(oInp)
      End If
    End If
  End Function

  Public Function FixParam(ByVal oIm As Object, ByVal bIsNullable As Boolean) As String
    If CNullS(oIm) = "" Then
      FixParam = ""
    ElseIf TypeOf oIm Is String Then
      FixParam = "'" & Trim(Replace(oIm, "'", "''")) & "'"
    ElseIf TypeOf oIm Is Boolean Then
      If CBool(oIm) Then
        FixParam = "True"
      Else
        FixParam = "False"
      End If
    Else
      FixParam = CNullS(oIm)
    End If
    If FixParam.Trim.ToString = "" Then
      If bIsNullable Then
        FixParam = "Null"
      End If
    End If

  End Function


  Public Function PropertyAvailable(ByVal iPropertyID As Integer, ByVal dArriveDate As Date, ByVal dDepartDate As Date, ByRef sResult As String, Optional ByVal bCheckAndSetLateCancellation As Boolean = False, Optional ByRef iLateCancelBookID As Integer = 0, Optional ByVal iLateCancelRebookID As Integer = 0, Optional ByRef sWhereUsed As String = "") As Boolean
    PropertyAvailable = False

    Dim oHouse As HouseRentalDataContext = MyData.GetHouseRentalContext(True)
    Try

      Dim oBook As Booking = Nothing

      If iPropertyID > 0 Then

        If bCheckAndSetLateCancellation Then
          oBook = (From x In oHouse.Bookings Where x.Status = "Cancelled" And x.LateCancellationRefundStatus = "None" And x.Property_ID = iPropertyID And x.ArriveDate >= dArriveDate And x.DepartDate <= dDepartDate).FirstOrDefault
        Else
          oBook = (From x In oHouse.Bookings Where x.Status <> "Cancelled" And x.Property_ID = iPropertyID And x.ArriveDate >= dArriveDate And x.DepartDate <= dDepartDate).FirstOrDefault
        End If

        If oBook Is Nothing Then
          PropertyAvailable = True
        Else
          If bCheckAndSetLateCancellation Then
            Dim cLateCancelFee As Long = 0, dRefundAmt As Long = 0
            iLateCancelBookID = oBook.Booking_ID
            oBook.LateCancellationRefundReBookingID = iLateCancelRebookID
            ' BJR the following is hard coded.  Could be pulled from CodeData table, but hasn't changed in years
            oBook.LateCancellationFeeAmount = oBook.Rate * 0.1
            oBook.LateCancellationRefundStatus = "In Process"
            oBook.CancellationHostPaymentProcessed = Nothing
            oHouse.SubmitChanges()

            PropertyAvailable = True
          Else
            sResult = "We're sorry, but the property you initially chose<BR/>is now unavailable.<BR/>Please select another property.<BR/>"
          End If
        End If
      Else
        sResult = "There was a problem getting the property information.<BR/>Please try again later.<BR/>"
        oErr.LogWebError("Property ID was not available when checking for Property Availability", "Reserve.aspx:PropertyAvailable", True)
      End If
      oBook = Nothing
    Catch ex As Exception
      oErr.HandleError(ex)
    End Try
    MyData.DisposeHouseRentalContext(oHouse)
  End Function


  Public Function BrowserInCStayOffice(ByRef oRequest As System.Web.HttpRequest, ByRef oSession As System.Web.SessionState.HttpSessionState) As Boolean

    If CNullS(oSession("LonettaOffice")) = "" Then

      Try
        Dim oCode = (From x In MyData.HouseRentalContext.CodeDatas Where x.CodeKey = "LonettaHostName").FirstOrDefault
        If oCode IsNot Nothing Then
          Dim oIP As IPHostEntry = Nothing
          oIP = Dns.GetHostEntry(CNullS(oCode.CodeValue))
          If oRequest.UserHostAddress = oIP.AddressList(0).ToString Then
            oSession("LonettaOffice") = "True"
          Else
            oSession("LonettaOffice") = "False"
          End If
        End If
      Catch ex As Exception

      End Try
    End If
    BrowserInCStayOffice = oSession("LonettaOffice") = "True"

  End Function

  Public Shared Function GetAppSetting(ByVal sKey As String, Optional ByVal sDefault As String = "") As String
    ' Only for Reserve website, not useable from version compiled for COM access by Rental SW
    '    GetAppSetting = IIf(IsNothing(System.Configuration.ConfigurationManager.AppSettings(sKey)), sDefault, System.Configuration.ConfigurationManager.AppSettings(sKey))
    GetAppSetting = sDefault
  End Function


#End Region


End Class
