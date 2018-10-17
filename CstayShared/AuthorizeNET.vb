Imports System.Linq
Imports System.Data.Linq
Imports AuthorizeNet.Api.Controllers
Imports AuthorizeNet.Api.Contracts.V1
Imports AuthorizeNet.Api.Controllers.Bases


Public Class CStayAuthorizeNET

  Public oErr As ErrorHandler
  Public oShared As SharedUtilities

  Public sSystemEmailAddress As String = ""
  Public sConnString As String = ""

  Public PaymentMethod As String
  Public InpLoginID As String = ""
  Public InpTransactionKey As String = ""
  Public InpTestMode As Boolean = False
  Public InpDelimiter As String = "|"
  Public InpRelayResponse As Boolean = False
  Public InpChargeAmount As Double = 0
  Public InpAuthorizationType As String = "AUTH_CAPTURE"
  Public InpPrevTransactionID As String = ""
  Public InpVersion As String = "3.1"
  Public InpCardNumber As String = ""
  Public InpExpMonth As String = ""
  Public InpExpYear As String = ""
  Public InpVerficationCode As String = ""
  Public InpFirstName As String = ""
  Public InpLastName As String = ""
  Public InpAddress As String = ""
  Public InpCity As String = ""
  Public InpState As String = ""
  Public InpZip As String = ""
  Public InpCustomerID As String = ""
  Public InpInvoiceNumber As String = ""
  Public PostURL As String = ""
  Public RawResult As String = ""
  Public CimLibKey As String = "MB7M-8YLC-YVWH-DHPM-BAXH"
  Public CimLibKeyName As String = "Byron Richard"
  Public CimLibKeyEmail As String = "shop@trinitysoftware.net"

  Public InpRecurringBilling As String = "NO"
  Public InpBankABANumber As String = ""
  Public InpBankAccountNumber As String = ""
  Public InpBankAccountType As String = ""
  Public InpBankName As String = ""
  Public InpBankAccountName As String = ""
  Public InpeCheckType As String = "WEB"


  Public RespCode As String = ""
  Public RespSubcode As String = ""
  Public RespReasonCode As String = ""
  Public RespReasonText As String = ""
  Public UserRespReasonText As String = ""
  Public RespApprovalCode As String = ""
  Public RespAVSCode As String = ""
  Public RespAVSText As String = ""
  Public UserRespAVSText As String = ""
  Public RespTransID As String = ""
  Public RespCVCCCode As String = ""
  Public RespCVCCText As String = ""

  Public CStayChargeResultText As String = ""
  Public UserChargeResultText As String = ""
  Public TransactionBeingHeld As Boolean = False
  '  Public CancelHeldTransaction As Boolean = True

  Public Sub Clear()
    InpLoginID = ""
    InpTransactionKey = ""
    InpTestMode = False
    InpDelimiter = "|"
    InpRelayResponse = False
    InpChargeAmount = 0
    InpAuthorizationType = "AUTH_CAPTURE"
    InpCardNumber = ""
    InpExpMonth = ""
    InpExpYear = ""
    InpVerficationCode = ""
    InpFirstName = ""
    InpLastName = ""
    InpAddress = ""
    InpCity = ""
    InpState = ""
    InpZip = ""
    InpCustomerID = ""
    PostURL = ""
    RawResult = ""

    RespCode = ""
    RespSubcode = ""
    RespReasonCode = ""
    RespReasonText = ""
    RespApprovalCode = ""
    RespAVSCode = ""
    RespAVSText = ""
    RespTransID = ""
    RespCVCCCode = ""
    RespCVCCText = ""
    UserChargeResultText = ""
    UserRespAVSText = ""
    UserRespReasonText = ""

    CStayChargeResultText = ""
    InpBankABANumber = ""
    InpBankAccountNumber = ""
    InpBankAccountType = ""
    InpBankName = ""
    InpBankAccountName = ""

    PaymentMethod = ""

  End Sub

  Public Function CreateURL() As Boolean

    CreateURL = False
    Try

      PostURL = "x_login=" & InpLoginID
      PostURL += "&x_tran_key=" & InpTransactionKey
      PostURL += "&x_version=3.1"
      If InpTestMode Then
        PostURL += "&x_test_request=TRUE"
      End If
      If InpDelimiter.ToString <> "" Then
        PostURL += "&x_delim_data=TRUE"
        PostURL += "&x_delim_char=" & InpDelimiter
      End If

      PostURL += "&x_relay_response=" & InpRelayResponse.ToString
      PostURL += "&x_amount=" & System.Web.HttpUtility.UrlEncode(InpChargeAmount.ToString)
      If InpInvoiceNumber.ToString <> "" Then PostURL += "&x_invoice_num=" & System.Web.HttpUtility.UrlEncode(InpInvoiceNumber)
      If InpCustomerID.ToString <> "" Then PostURL += "&x_cust_id=" & System.Web.HttpUtility.UrlEncode(InpCustomerID)
      If InpFirstName.ToString <> "" Then PostURL += "&x_first_name=" & System.Web.HttpUtility.UrlEncode(InpFirstName)
      If InpLastName.ToString <> "" Then PostURL += "&x_last_name=" & System.Web.HttpUtility.UrlEncode(InpLastName)
      If InpAddress.ToString <> "" Then PostURL += "&x_address=" & System.Web.HttpUtility.UrlEncode(InpAddress)
      If InpCity.ToString <> "" Then PostURL += "&x_city=" & System.Web.HttpUtility.UrlEncode(InpCity)
      If InpState.ToString <> "" Then PostURL += "&x_state=" & System.Web.HttpUtility.UrlEncode(InpState)
      If InpZip.ToString <> "" Then PostURL += "&x_zip=" & System.Web.HttpUtility.UrlEncode(InpZip)
      If PaymentMethod.ToString = "Card" Then
        PostURL += "&x_method=CC"
        If InpAuthorizationType = "CREDIT" Or InpAuthorizationType = "VOID" Then
          If InpPrevTransactionID <> "" Then
            PostURL += "&x_trans_id=" & System.Web.HttpUtility.UrlEncode(InpPrevTransactionID)
          End If
        End If
        If InpVerficationCode <> "" Then PostURL += "&x_card_code=" & CStr(InpVerficationCode)
        If InpCardNumber.ToString <> "" Then PostURL += "&x_card_num=" & InpCardNumber
        If InpExpMonth.ToString <> "" And InpExpYear.ToString <> "" Then PostURL += "&x_exp_date=" & System.Web.HttpUtility.UrlEncode(InpExpMonth & InpExpYear)
      Else
        PostURL += "&x_method=ECHECK"
        If InpBankABANumber.ToString <> "" Then PostURL += "&x_bank_aba_code=" & System.Web.HttpUtility.UrlEncode(InpBankABANumber)
        If InpBankAccountName.ToString <> "" Then PostURL += "&x_bank_acct_name=" & System.Web.HttpUtility.UrlEncode(InpBankAccountName)
        If InpBankAccountNumber.ToString <> "" Then PostURL += "&x_bank_acct_num=" & System.Web.HttpUtility.UrlEncode(InpBankAccountNumber)
        If InpBankAccountType.ToString <> "" Then PostURL += "&x_bank_acct_type=" & System.Web.HttpUtility.UrlEncode(InpBankAccountType)
        If InpBankName.ToString <> "" Then PostURL += "&x_bank_name=" & System.Web.HttpUtility.UrlEncode(InpBankName)
        PostURL += "&x_echeck_type=" & System.Web.HttpUtility.UrlEncode(InpeCheckType)
      End If
      PostURL += "&x_type=" & InpAuthorizationType



      PostURL = PostURL
      CreateURL = True
    Catch ex As Exception
      oErr.HandleError(ex)
    End Try


  End Function

  Function MakeCharge(bChargeInTestMode As Boolean, ByRef sErr As String) As Boolean

    MakeCharge = False
    Try
      If PostURL.ToString <> "" Then
        RawResult = RetrieveHtmlPage(PostURL, bChargeInTestMode, sErr)
      End If
      MakeCharge = UnwrapChargeResult()

    Catch ex As Exception
      oErr.HandleError(ex)
    End Try

  End Function



  Function UnwrapChargeResult() As Boolean

    UnwrapChargeResult = False
    TransactionBeingHeld = False
    Try

      If RawResult.ToString <> "" Then
        Dim sData() As String
        sData = Split(RawResult, InpDelimiter)
        If UBound(sData) > 6 Then
          RespCode = sData(0)
          RespSubcode = sData(1)
          RespReasonCode = sData(2)
          RespReasonText = sData(3)
          RespApprovalCode = sData(4)
          RespAVSCode = sData(5)
          GetAVSReasonText(RespAVSCode)
          RespTransID = sData(6)
        End If
      End If

      If RespCode.ToString = "1" Then
        UnwrapChargeResult = True
      ElseIf RespCode.ToString = "2" Then
        UserChargeResultText = "Our attempt to charge your credit card has been declined."
        If RespReasonCode = "27" Then
          UserChargeResultText &= "<br>The Address Verification System gave this reason: " & UserRespAVSText
        Else
          UserChargeResultText &= GetUserDenialReasonText(RespReasonCode)
        End If
        UserChargeResultText &= "<br>Please check your card information and try again or call <span class=""SmallTitle"">607-547-6260</span> .<BR/>"
        CStayChargeResultText = "The charge was declined.  The reason given was: " & GetResponseReasonDetails(RespReasonCode, RespReasonText)
        If RespAVSText <> "" Then CStayChargeResultText &= vbCrLf & "The AVS Result was: " & RespAVSText

      ElseIf RespCode.ToString = "3" Then
        CStayChargeResultText = "There was an error while making the charge.  The reason given was: " & GetResponseReasonDetails(RespReasonCode, RespReasonText)
        If RespReasonCode = "11" Then
          UserChargeResultText = "A duplicate credit card charge has been detected.<BR/>You may have accidentally clicked the 'Submit' button twice.<BR/>Please call <span class=""SmallTitle"">607-547-6260</span> to verify your payment."
        Else
          UserChargeResultText = "We're sorry, an error ocurred while processing your credit card.<br>Please check your card information and try again or call <span class=""SmallTitle"">607-547-6260</span> .<BR/>"
        End If
      ElseIf RespCode.ToString = "4" Then
        CStayChargeResultText = "The credit card Transaction with Authorize.NET TransactionID=" & RespTransID & " is being held for merchant review.  Please follow up on the transaction to see if it clears.  If it does not, you will need to contact the customer to get another card number. "
        UserChargeResultText = "<span style=""font-family:Courier New; font-size:14px; font-weight:bold; color:Red"" > IMPORTANT: Our cedit card processor is holding this charge for review.<br>If the charge does not go through, we will contact you to obtain another means of payment.<br></span>"
        oErr.LogWebError(CStayChargeResultText)
        UnwrapChargeResult = True
        TransactionBeingHeld = True
      End If


    Catch ex As Exception
      oErr.HandleError(ex)
    End Try

  End Function

  Public Function GetResponseReasonDetails(ByVal sCode As String, ByRef sReasonText As String) As String

    GetResponseReasonDetails = sReasonText
    Select Case sCode
      Case "4"
        GetResponseReasonDetails &= "  The card needs picked up"
      Case "44"
        GetResponseReasonDetails &= "  The card code submitted with the transaction did not match the card code on file at the card issuing bank and the transaction was declined."
      Case "44"
        GetResponseReasonDetails &= "  The Merchant Interface is set to reject transactions with certain values for a Card Code mismatch."
      Case "103"
        GetResponseReasonDetails &= "  A valid fingerprint, Transaction Key, or password is required for this transaction."
      Case "120"
        GetResponseReasonDetails &= "  The original transaction timed out while waiting for a response from the authorizer."
      Case "121"
        GetResponseReasonDetails &= "  The original transaction experienced a database error."
      Case "122"
        GetResponseReasonDetails &= "  The original transaction experienced a processing error."
      Case "128"
        GetResponseReasonDetails &= "  The customer’s financial institution does not currently allow transactions for this account."
      Case "250"
        GetResponseReasonDetails &= "  This transaction was submitted from a blocked IP address."
      Case "251"
        GetResponseReasonDetails &= "  The transaction was declined as a result of triggering a Fraud Detection Suite filter."
      Case "252", "253"
        GetResponseReasonDetails &= "  The transaction was accepted, but is being held for merchant review"
    End Select
  End Function


  Public Function GetUserDenialReasonText(ByVal sCode As String) As String

    GetUserDenialReasonText = ""
    Select Case sCode
      Case "5"
        GetUserDenialReasonText = "The charge amount is invalid."
      Case "6"
        GetUserDenialReasonText = "Invalid credit card number."
      Case "7"
        GetUserDenialReasonText = "Invalid card expiration date."
      Case "8"
        GetUserDenialReasonText = "Credit card has expired."
      Case "17"
        GetUserDenialReasonText = "Card type not accepted."
    End Select
    If GetUserDenialReasonText <> "" Then
      UserRespReasonText = GetUserDenialReasonText
      GetUserDenialReasonText = "The credit card processor gave this reason: " & GetUserDenialReasonText & "<br>"
    End If

  End Function
  Public Function GetCVCCReasonText(ByVal sCode As String) As String

    GetCVCCReasonText = ""
    Select Case sCode
      Case "M"
        GetCVCCReasonText = "Verification code matches card"
      Case "N"
        GetCVCCReasonText = "Verification code does not match card"
      Case "P"
        GetCVCCReasonText = "Not processed"
      Case "S"
        GetCVCCReasonText = "Card should have been present"
      Case "U"
        GetCVCCReasonText = "Card issues unable to process request"
    End Select
  End Function
  Public Sub GetAVSReasonText(ByVal sCode As String)

    RespAVSText = ""
    Select Case sCode
      Case "A"
        RespAVSText = "Address matches card, but zip code does not"
      Case "B"
        RespAVSText = "Address information was missing"
      Case "E"
        RespAVSText = "Error while processing"
        UserChargeResultText = "There was an error while Verifying your Address."
      Case "G"
        RespAVSText = "Card is from a non-U.S. bank"
      Case "N"
        RespAVSText = "Address and zip do not match with card"
      Case "P"
        RespAVSText = "AVS not applicable for this transaction"
      Case "R"
        RespAVSText = "System unavailable"
        UserRespAVSText = "The Address Verification System was not available."
      Case "S"
        RespAVSText = "Card issuer does not support AVS"
        UserRespAVSText = "Your card issuer does not support Address Verification."
      Case "U"
        RespAVSText = "Address information was not available"
      Case "W"
        RespAVSText = "9 digit zip matches, card but address does not"
      Case "X"
        RespAVSText = "Address and 9 digit zip do not match"
      Case "Y"
        RespAVSText = "Address and 5 digit zip do not match"
      Case "Z"
        RespAVSText = "5 digit zip matches card, but address does not"
    End Select
    If UserRespAVSText = "" Then UserRespAVSText = RespAVSText

  End Sub

  Public Function RetrieveHtmlPage(ByVal sRequest As String, bChargeInTestMode As Boolean, ByRef sErr As String) As String

    RetrieveHtmlPage = ""


    Dim myWriter As System.IO.StreamWriter = Nothing

    System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
    Dim objRequest As System.Net.HttpWebRequest
    If bChargeInTestMode Then
      objRequest = CType(System.Net.WebRequest.Create("https://test.authorize.net/gateway/transact.dll?"), System.Net.HttpWebRequest)
    Else
      objRequest = CType(System.Net.WebRequest.Create("https://secure2.authorize.net/gateway/transact.dll?"), System.Net.HttpWebRequest)
    End If
    objRequest.Method = "POST"
    objRequest.ContentLength = sRequest.Length
    objRequest.ContentType = "application/x-www-form-sURLencoded"

    Try
      myWriter = New System.IO.StreamWriter(objRequest.GetRequestStream())
      myWriter.Write(sRequest)
    Catch ex As Exception
      oErr.HandleError(ex, False, True)
      sErr = oErr.GetError(ex)
    Finally
      myWriter.Close()

      Dim objResponse As System.Net.HttpWebResponse = CType(objRequest.GetResponse(), System.Net.HttpWebResponse)
      Dim myReader As New System.IO.StreamReader(objResponse.GetResponseStream())
      RetrieveHtmlPage = myReader.ReadToEnd()

      ' Close and clean up the StreamReader
      myReader.Close()
    End Try

  End Function

  '  Public Sub RecordCIMGatewayActivity(ByRef oCIMGate As TableCIMGatewayActivity, ByRef oCIM As ITDevWorks.CimLib.CimLib,
  '   Optional ByVal iBookingID As Integer = 0, Optional ByVal iGuestID As Integer = 0, Optional ByVal iPaymentID As Integer = 0, Optional ByVal sPaymentCategory As String = "")

  '    Try
  '      If oCIMGate Is Nothing Then oCIMGate = New TableCIMGatewayActivity
  '      oCIMGate.Clear()
  '      oCIMGate.BookingID__Integer = iBookingID
  '      oCIMGate.GatewayRequest__String = oCIM.LastResponse.RequestText
  '      oCIMGate.GatewayResponse__String = oCIM.LastResponse.ResponseXmlText
  '      oCIMGate.GatewayReturn__String = oCIM.LastResponse.ReturnDataXml
  '      If oCIM.LastResponse.ErrorsOccurred Then
  '        oCIMGate.GatewayError__String = oCIM.LastResponse.GetFullDescription
  '      End If
  '      oCIMGate.GuestID__Integer = iGuestID
  '      oCIMGate.PaymentCategory__String = sPaymentCategory
  '      oCIMGate.PaymentID__Integer = iPaymentID
  '      oCIMGate.Insert()
  '    Catch ex As Exception
  '      oErr.HandleError(ex, False)
  '    End Try

  '  End Sub

  '  Public Function GetCIMPaymentInformation(ByRef sResult As String, ByRef sCustomerCIMID As String, ByRef sPaymentCIMID As String,
  '  ByRef sCCPaymentID() As String, ByRef sCCNumber() As String, ByRef sCCFirstName() As String, ByRef sCCLastName() As String,
  '  ByRef sCCAddress() As String, ByRef sCCCity() As String, ByRef sCCState() As String, ByRef sCCZip() As String,
  '  Optional ByVal bTestMode As Boolean = False, Optional ByVal bGetAllPaymentsIfRequestedPaymentMissing As Boolean = False) As Boolean

  '    GetCIMPaymentInformation = False
  '    Try

  '      Dim oCIM As New ITDevWorks.CimLib.CimLib(CimLibKey, CimLibKeyName, CimLibKeyEmail), oCCust As ITDevWorks.CimLib.CustomerProfile = Nothing, oCPay As ITDevWorks.CimLib.PaymentProfile = Nothing
  '      Dim sData() As String = Nothing, iCount As Integer = 0

  '      If bTestMode Then
  '        oCIM.ApiLogin = "38vV43Dhg"
  '        oCIM.TransactionKey = "426bkKR4q6X2F97c"
  '        oCIM.GatewayUrl = ITDevWorks.CimLib.StandardGatewayUrl.Test
  '      Else
  '        oCIM.ApiLogin = "33WmD7pGbd"
  '        oCIM.TransactionKey = "3q3S4N9Lz2d7Nw3v"
  '        oCIM.GatewayUrl = ITDevWorks.CimLib.StandardGatewayUrl.Live
  '      End If

  '      If sCustomerCIMID = "" Then
  '        sResult = "No Customer ID was provided"
  '        Exit Function
  '      End If
  '      oCCust = oCIM.GetCustomerProfile(sCustomerCIMID)
  '      If oCIM.LastResponse.ErrorsOccurred Then
  '        If oCIM.LastResponse.ReasonCode = "E00040" Then
  '          sResult = "Customer record does not exist at Authorize.NET"
  '          Exit Function
  '        End If
  '      End If

  '      ReDim sCCPaymentID(0)
  '      ReDim sCCNumber(0)
  '      ReDim sCCFirstName(0)
  '      ReDim sCCLastName(0)
  '      ReDim sCCAddress(0)
  '      ReDim sCCCity(0)
  '      ReDim sCCState(0)
  '      ReDim sCCZip(0)

  '      If sPaymentCIMID <> "" Then
  '        oCPay = oCIM.GetPaymentProfile(sCustomerCIMID, sPaymentCIMID)
  '        If oCIM.LastResponse.ErrorsOccurred Then
  '          If oCIM.LastResponse.ReasonCode = "E00040" Then
  '            If bGetAllPaymentsIfRequestedPaymentMissing Then
  '              GoTo GetAllPayments
  '            Else
  '              sResult = "Payment record does not exist at Authorize.NET"
  '              Exit Function
  '            End If
  '          End If
  '        Else

  '          sCCPaymentID(0) = oCPay.Id
  '          sCCNumber(0) = Right(oCPay.CreditCardNumber, 4)
  '          sCCFirstName(0) = oCPay.FirstName
  '          sCCLastName(0) = oCPay.LastName
  '          sCCAddress(0) = oCPay.StreetAddress
  '          sCCCity(0) = oCPay.City
  '          sCCState(0) = oCPay.State
  '          sCCZip(0) = oCPay.Zip

  '        End If
  '      Else
  'GetAllPayments:
  '        If oCCust.PaymentProfiles.Count > 0 Then
  '          For Each oCThisPay As ITDevWorks.CimLib.PaymentProfile In oCCust.PaymentProfiles
  '            oCPay = oCThisPay

  '            ReDim Preserve sCCPaymentID(iCount)
  '            ReDim Preserve sCCNumber(iCount)
  '            ReDim Preserve sCCFirstName(iCount)
  '            ReDim Preserve sCCLastName(iCount)
  '            ReDim Preserve sCCAddress(iCount)
  '            ReDim Preserve sCCCity(iCount)
  '            ReDim Preserve sCCState(iCount)
  '            ReDim Preserve sCCZip(iCount)

  '            sCCPaymentID(iCount) = oCPay.Id
  '            sCCNumber(iCount) = Right(oCPay.CreditCardNumber, 4)
  '            sCCFirstName(iCount) = oCPay.FirstName
  '            sCCLastName(iCount) = oCPay.LastName
  '            sCCAddress(iCount) = oCPay.StreetAddress
  '            sCCCity(iCount) = oCPay.City
  '            sCCState(iCount) = oCPay.State
  '            sCCZip(iCount) = oCPay.Zip
  '            iCount += 1
  '          Next
  '        End If

  '      End If
  '      GetCIMPaymentInformation = True
  '    Catch ex As Exception
  '      oErr.HandleError(ex, False)
  '    End Try

  '  End Function


  '  Public Function CreateCIMPaymentInformation(ByRef sResult As String, ByVal sGuestID As String, ByVal sGuestName As String,
  '   ByVal sGuestEmail As String, ByRef sCustomerCIMID As String, ByRef sPaymentCIMID() As String, ByRef sPaymentID() As String, ByRef sPaymentCategory() As String,
  '   ByRef sCCNumber() As String, ByRef sCCMonth() As String, ByRef sCCYear() As String, ByRef sCCFirstName() As String, ByRef sCCLastName() As String,
  '   ByRef sCCAddress() As String, ByRef sCCCity() As String, ByRef sCCState() As String, ByRef sCCZip() As String,
  '   Optional ByVal bTestMode As Boolean = False) As Boolean

  '    CreateCIMPaymentInformation = False
  '    Dim sOut As String = ""
  '    Dim oCIM As New ITDevWorks.CimLib.CimLib(CimLibKey, CimLibKeyName, CimLibKeyEmail), oCCust As ITDevWorks.CimLib.CustomerProfile = Nothing, oCPay As ITDevWorks.CimLib.PaymentProfile = Nothing
  '    Dim sData() As String = Nothing, iCount As Integer = 0

  '    Try


  '      If bTestMode Then
  '        oCIM.ApiLogin = "38vV43Dhg"
  '        oCIM.TransactionKey = "426bkKR4q6X2F97c"
  '        oCIM.GatewayUrl = ITDevWorks.CimLib.StandardGatewayUrl.Test
  '      Else
  '        oCIM.ApiLogin = "33WmD7pGbd"
  '        oCIM.TransactionKey = "3q3S4N9Lz2d7Nw3v"
  '        oCIM.GatewayUrl = ITDevWorks.CimLib.StandardGatewayUrl.Live
  '      End If

  '      If sGuestID = "" Then
  '        sResult = "No Guest ID was provided"
  '        Exit Function
  '      End If
  '      If sGuestName = "" Then
  '        sResult = "No Guest Name was provided"
  '        Exit Function
  '      End If
  '      If sGuestEmail = "" Then
  '        sResult = "No Guest Email was provided"
  '        Exit Function
  '      End If

  '      If sCustomerCIMID <> "" Then
  '        oCCust = oCIM.GetCustomerProfile(sCustomerCIMID)
  '        If oCIM.LastResponse.ErrorsOccurred Then
  '          If oCIM.LastResponse.ReasonCode = "E00040" Then
  '            Throw New Exception("Could not retrieve Authorize.NET Customer Info for GuestID=" & sGuestID & ", CustomerID=" & sCustomerCIMID & vbCrLf & oCIM.LastResponse.ReasonCode & vbCrLf & oCIM.LastResponse.GetFullDescription)
  '          End If
  '        End If
  '      Else
  '        oCCust = oCIM.CreateCustomerProfile(sGuestID)
  '        oCCust.ReferenceId = sGuestID
  '        oCCust.Description = sGuestName & "-" & sGuestEmail
  '        oCIM.CustomerProfiles.Add(oCCust)
  '        If oCIM.LastResponse.ErrorsOccurred Then
  '          If oCIM.LastResponse.ReasonCode = "E00039" Then
  '            sData = Split(oCIM.LastResponse.ReasonText, "id ")
  '            sData = Split(sData(UBound(sData)), " ")
  '            If UBound(sData) > 0 Then
  '              sCustomerCIMID = sData(0)
  '            End If
  '          End If
  '          If sCustomerCIMID = "" Then Throw New Exception("Could not create Customer through the Authorize.NET CIM for GuestID=" & sGuestID & vbCrLf & oCIM.LastResponse.ReasonCode & vbCrLf & oCIM.LastResponse.GetFullDescription)
  '          oCCust = oCIM.GetCustomerProfile(sCustomerCIMID)
  '        End If
  '        sCustomerCIMID = oCCust.Id
  '      End If

  '      ReDim sPaymentCIMID(UBound(sCCNumber))

  '      For iCount = 0 To UBound(sCCNumber)

  '        If sCCNumber(iCount) <> "" Then
  '          oCPay = oCIM.CreatePaymentProfile()
  '          If sCCMonth(iCount).Trim.Length = 1 Then sCCMonth(iCount) = "0" & sCCMonth(iCount)
  '          sCCYear(iCount) = sCCYear(iCount).Substring(sCCYear(iCount).Length - 2, 2)
  '          oCPay.CreditCardNumber = sCCNumber(iCount)
  '          oCPay.CreditCardExpiration = CDate(sCCMonth(iCount) & "/1/" & sCCYear(iCount))
  '          oCPay.FirstName = sCCFirstName(iCount)
  '          oCPay.LastName = sCCLastName(iCount)
  '          oCPay.StreetAddress = sCCAddress(iCount)
  '          oCPay.City = sCCCity(iCount)
  '          oCPay.State = sCCState(iCount)
  '          oCPay.Zip = sCCZip(iCount)
  '          oCPay.ReferenceId = sPaymentID(iCount) & ":" & sPaymentCategory(iCount)
  '          oCCust.PaymentProfiles.Add(oCPay)
  '          If oCIM.LastResponse.ErrorsOccurred Then
  '            If sCustomerCIMID = "" Then Throw New Exception("Could not create Payment record through the Authorize.NET CIM for GuestID=" & sGuestID & ", PaymentID=" & sPaymentID(iCount) & vbCrLf & oCIM.LastResponse.ReasonCode & vbCrLf & oCIM.LastResponse.GetFullDescription)
  '          Else
  '            sPaymentCIMID(iCount) = oCPay.Id
  '          End If


  '        End If

  '      Next

  '      CreateCIMPaymentInformation = True
  '    Catch ex As Exception

  '      sResult &= "ERROR:" & vbCrLf & "Message: " & ex.Message.ToString & vbCrLf & "Site:  " & ex.TargetSite.ToString & vbCrLf & "Stack Trace:  " & ex.StackTrace.ToString & vbCrLf & "Input Vars: " & vbCrLf & sOut
  '      oErr.HandleError(ex, False)
  '    End Try

  '  End Function


  Public Function ProcessCharge(ByRef oAuth As CStayAuthorizeNET, ByRef oBook As TableBookings, ByRef oPayment As Object, ByRef oGuest As TableGuests, ByRef oProp As TableProperties, ByRef oCIMGate As TableCIMGatewayActivity,
  ByRef oHost As TableHosts, ByVal dAmountCharged As Double, ByVal sPaymentOrigin As String, ByVal sPaymentMethod As String,
  ByVal sPaymentLocation As String, ByRef sCustomerCIMID As String, ByRef sPaymentCIMID As String, ByVal sCCNumber As String, ByVal sCCVerif As String, ByVal sCCMonth As String, ByVal sCCYear As String,
  ByVal sCCFirstName As String, ByVal sCCLastName As String, ByVal sCCAddress As String, ByVal sCCCity As String,
  ByVal sCCState As String, ByVal sCCZip As String, ByVal sBankABANumber As String, ByVal sBankAccountName As String,
  ByVal sBankAccountNumber As String, ByVal sBankAccountType As String, ByVal sBankName As String, ByRef sCSTAYResult As String, ByRef sUserResult As String,
  ByVal bChargeInTestMode As Boolean, ByVal bPutErrorInResult As Boolean, ByVal bSkipMakingCharge As Boolean,
  Optional ByVal sCreditType As String = "", Optional ByVal sPrevTransactionID As String = "", Optional ByRef sPostURLUsed As String = "",
  Optional ByVal bUseCIMGateway As Boolean = True, Optional ByRef bCIMTransactionDeclined As Boolean = False, Optional ByRef bProblemCard As Boolean = False) As Boolean

    ProcessCharge = False
    oAuth.Clear()
    Dim iTryNumber = 0

    If bUseCIMGateway Then
      Try

TryAgain:
        Dim iBookingID As Integer = oBook.Booking_ID_PK__Integer, iGuestID As Integer = oGuest.Guest_ID_PK__Integer, iPaymentID As Integer = oPayment.PaymentID_PK__Integer
        '        Dim oCIM As New ITDevWorks.CimLib.CimLib(CimLibKey, CimLibKeyName, CimLibKeyEmail), oCCust As ITDevWorks.CimLib.CustomerProfile = Nothing, oCPay As ITDevWorks.CimLib.PaymentProfile = Nothing
        Dim sCustRefID As String = Now.ToString.Replace(" ", "").Replace(":", "").Replace("/", ""), sPayRefID As String = Now.ToString.Replace(" ", "").Replace(":", "").Replace("/", "")
        Dim bCreateCustomer As Boolean = False, bAddPayment As Boolean = False, iIndex As Integer = 0
        Dim sAPILogin As String = "", sTransactionKey As String = "", sGatewayURL As String = "", sErrMsg As String = ""

        If sCCMonth.Trim.Length = 1 Then sCCMonth = "0" & sCCMonth
        sCCYear = sCCYear.Substring(sCCYear.Length - 2, 2)

        If bChargeInTestMode Then
          ApiOperationBase(Of ANetApiRequest, ANetApiResponse).RunEnvironment = AuthorizeNet.Environment.SANDBOX
          ApiOperationBase(Of ANetApiRequest, ANetApiResponse).MerchantAuthentication = New merchantAuthenticationType() With {.name = "38vV43Dhg", .ItemElementName = ItemChoiceType.transactionKey, .Item = "426bkKR4q6X2F97c"}
        Else
          ApiOperationBase(Of ANetApiRequest, ANetApiResponse).RunEnvironment = AuthorizeNet.Environment.PRODUCTION
          ApiOperationBase(Of ANetApiRequest, ANetApiResponse).MerchantAuthentication = New merchantAuthenticationType() With {.name = "33WmD7pGbd", .ItemElementName = ItemChoiceType.transactionKey, .Item = "3q3S4N9Lz2d7Nw3v"}
        End If

        If sCustomerCIMID = "" Then
          bCreateCustomer = True
        Else
          Dim request = New getCustomerProfileRequest()
          request.customerProfileId = sCustomerCIMID
          Dim controller = New getCustomerProfileController(request)
          controller.Execute()
          Dim response = controller.GetApiResponse()

          If response Is Nothing Or response.messages.resultCode <> messageTypeEnum.Ok Then
            bCreateCustomer = True
          Else
            If sPaymentCIMID = "" Then

              If response.profile.paymentProfiles IsNot Nothing AndAlso response.profile.paymentProfiles.Length > 0 Then
                For i As Integer = 0 To response.profile.paymentProfiles.Length - 1
                  If ((TryCast(response.profile.paymentProfiles(i).payment.Item, creditCardMaskedType)).cardNumber.ToString().Contains(Right(sCCNumber, 4))) Then
                    sPaymentCIMID = response.profile.paymentProfiles(i).customerPaymentProfileId
                  End If
                Next
              End If
              If sPaymentCIMID = "" Then bAddPayment = True
            End If
          End If
        End If

        Dim paymentProfileList As List(Of customerPaymentProfileType)
        If bCreateCustomer Or bAddPayment Then

          Dim billToAddress As customerAddressType = New customerAddressType()
          billToAddress.firstName = sCCFirstName
          billToAddress.lastName = sCCLastName
          billToAddress.address = sCCAddress
          billToAddress.city = sCCCity
          billToAddress.state = sCCState
          billToAddress.zip = sCCZip

          Dim creditCard = New creditCardType With {.cardNumber = sCCNumber, .expirationDate = sCCMonth & sCCYear}
          Dim cc As paymentType = New paymentType With {.Item = creditCard}
          Dim ccPaymentProfile As customerPaymentProfileType = New customerPaymentProfileType()
          ccPaymentProfile.payment = cc
          ccPaymentProfile.billTo = billToAddress

          If bCreateCustomer Then
            paymentProfileList = New List(Of customerPaymentProfileType)()
            paymentProfileList.Add(ccPaymentProfile)
          Else
            Dim request = New createCustomerPaymentProfileRequest With {.customerProfileId = sCustomerCIMID, .paymentProfile = ccPaymentProfile, .validationMode = validationModeEnum.none}
            Dim controller = New createCustomerPaymentProfileController(request)
            controller.Execute()
            Dim response As createCustomerPaymentProfileResponse = controller.GetApiResponse()

            If response IsNot Nothing Then
              If response.messages IsNot Nothing Then
                If response.messages.resultCode = messageTypeEnum.Ok Then
                  sPaymentCIMID = response.customerPaymentProfileId
                Else
                  If controller.GetErrorResponse().messages.message.Length > 0 Then
                    sErrMsg = "Could not create Payment Profile for CIMID = " & sCustomerCIMID & " Code = " & response.messages.message(0).code & vbCrLf & " Text= " & response.messages.message(0).text
                  Else
                    sErrMsg = "Could not create Payment Profile for CIMID = " & sCustomerCIMID & " Code: " & response.messages.resultCode & vbCrLf & response.messages.message.ToString()
                  End If
                End If
              Else
                sErrMsg = "Could not create Payment Profile for CIMID = " & sCustomerCIMID & ", Unknown Reason"
              End If
            Else
              sErrMsg = "Could not create Payment Profile for CIMID = " & sCustomerCIMID & ", No response"
            End If
          End If
          If sErrMsg <> "" Then
            Throw New Exception(sErrMsg)
          End If

        End If

        If bCreateCustomer Then

          Dim customerProfile As customerProfileType = New customerProfileType()
          customerProfile.merchantCustomerId = iBookingID
          customerProfile.email = oGuest.Email__String
          customerProfile.description = iBookingID & " - " & oGuest.GuestFirstName__String & " " & oGuest.GuestLastName__String & ", " & oGuest.Address__String & ", " & oGuest.City__String
          customerProfile.paymentProfiles = paymentProfileList.ToArray()
          Dim request = New createCustomerProfileRequest With {.profile = customerProfile, .validationMode = validationModeEnum.none}
          Dim controller = New createCustomerProfileController(request)
          controller.Execute()

          Dim response As createCustomerProfileResponse = controller.GetApiResponse()
          If response IsNot Nothing Then
            If response.messages.message IsNot Nothing Then
              If response.messages.resultCode = messageTypeEnum.Ok Then
                sCustomerCIMID = response.customerProfileId
                sPaymentCIMID = response.customerPaymentProfileIdList(0)
                oGuest.Clear()
                oGuest.Guest_ID_PK__Integer = oBook.Guest_ID__Integer
                oGuest.AuthNETCustomerProfileID__String = sCustomerCIMID
                oGuest.Update()
                oGuest.Clear()
                oGuest.Guest_ID_PK__Integer = oBook.Guest_ID__Integer
                oGuest.SelectData()
              End If
            Else
              If controller.GetErrorResponse().messages.message.Length > 0 Then
                sErrMsg = "Could not add new Authorize.NET Customer record for GuestID=" & oGuest.Guest_ID_PK__Integer & vbCrLf & " Code = " & response.messages.message(0).code & vbCrLf & " Text= " & response.messages.message(0).text
              Else
                sErrMsg = "Could not add new Authorize.NET Customer record for GuestID=" & oGuest.Guest_ID_PK__Integer & vbCrLf & " Code = " & response.messages.message(0).code & vbCrLf & " Text= " & response.messages.message(0).text
              End If
            End If
          Else
            sErrMsg = "Could not add new Authorize.NET Customer record for GuestID=" & oGuest.Guest_ID_PK__Integer & vbCrLf & " Cause Unknown "
          End If
          If sErrMsg <> "" Then
            Throw New Exception(sErrMsg)
          End If

        End If

        If bSkipMakingCharge Then
          RespApprovalCode = "Skip"
          RespCode = "Skip"
          RespReasonCode = "Skip"
          RespReasonText = "Skip"
          RespApprovalCode = "Skip"
          RespTransID = "Skip"
          ProcessCharge = True
        Else

          oPayment.Clear()
          oPayment.PaymentID_PK__Integer = iPaymentID
          oPayment.AuthNETPaymentProfileID__String = sPaymentCIMID
          oPayment.Update()
          oPayment.Clear()
          oPayment.PaymentID_PK__Integer = iPaymentID
          oPayment.SelectData()

          Dim ddAmountCharged As Decimal = dAmountCharged
          Dim sReservInfo As String = Left(Left(sPaymentLocation, 1) & Left(sPaymentOrigin, 1) & "-" & oBook.Booking_ID_PK__Integer & "-" & oHost.HostName__String, 20)

          Dim profileToCharge As customerProfilePaymentType = New customerProfilePaymentType()
          profileToCharge.customerProfileId = sCustomerCIMID
          profileToCharge.paymentProfile = New paymentProfile With {.paymentProfileId = sPaymentCIMID}
          Dim transOrder As New orderType
          transOrder.invoiceNumber = sReservInfo

          Dim transactionType As String = ""
          Dim transactionRequest = Nothing

          If sCreditType = "Refund" Then
            transactionType = transactionTypeEnum.refundTransaction.ToString()
            transactionRequest = New transactionRequestType With {.transactionType = transactionType, .amount = dAmountCharged, .profile = profileToCharge, .refTransId = sPrevTransactionID}
          ElseIf sCreditType = "Void" Then
            transactionType = transactionTypeEnum.voidTransaction.ToString()
            transactionRequest = New transactionRequestType With {.transactionType = transactionType, .refTransId = sPrevTransactionID}
          Else
            transactionType = transactionTypeEnum.authCaptureTransaction.ToString()
            transactionRequest = New transactionRequestType With {.transactionType = transactionType, .amount = dAmountCharged, .profile = profileToCharge, .order = transOrder}
          End If

          Dim request = New createTransactionRequest With {.transactionRequest = transactionRequest}
          Dim controller = New createTransactionController(request)
          controller.Execute()
          Dim response = controller.GetApiResponse()
          If response IsNot Nothing Then
            If response.messages.resultCode = messageTypeEnum.Ok Then
              If response.transactionResponse.messages IsNot Nothing Then
                RespTransID = response.transactionResponse.transId
                RespCode = response.transactionResponse.responseCode
                RespReasonCode = response.transactionResponse.messages(0).code
                RespReasonText = response.transactionResponse.messages(0).description
                RespApprovalCode = response.transactionResponse.authCode
                bCIMTransactionDeclined = False
              Else
                bCIMTransactionDeclined = True
                If response.transactionResponse.errors IsNot Nothing Then
                  RespReasonText = response.transactionResponse.errors(0).errorCode & " - " &
                                  response.transactionResponse.errors(0).errorText
                End If
              End If
            Else
              bCIMTransactionDeclined = True
              If response.transactionResponse IsNot Nothing AndAlso response.transactionResponse.errors IsNot Nothing Then
                RespReasonText = response.transactionResponse.errors(0).errorCode & " - " &
                                  response.transactionResponse.errors(0).errorText
              Else
                RespReasonText = response.messages.message(0).code & " - " & response.messages.message(0).text
              End If
            End If
          Else
            bCIMTransactionDeclined = True
          End If
          ProcessCharge = Not bCIMTransactionDeclined

          If bCIMTransactionDeclined Then
            UserChargeResultText = "We're sorry, an error ocurred while processing your credit card.<br>Please check your card information and try again or call <span class=""SmallTitle"">607-547-6260</span> .<BR/>"
            CStayChargeResultText = "Could not make payment through the Authorize.NET CIM for the following Reason: " & RespReasonText
            '            oErr.LogWebError("CIM Card Charge Declined: " & RespReasonText & "  :  BookingID=" & oBook.Booking_ID_PK__Integer & ", GuestID=" & oGuest.Guest_ID_PK__Integer, "AuthorizeNET:ProcessCharge")
          Else
            UnwrapChargeResult()
          End If
        End If

      Catch ex As Exception
        If ex.Message.ToString = "The request was aborted: Could not create SSL/TLS secure channel." And iTryNumber < 10 Then
          GoTo TryAgain
        End If
        iTryNumber = iTryNumber + 1
        If bPutErrorInResult Then sCSTAYResult = "ERROR:" & vbCrLf & "Message: " & ex.Message.ToString & vbCrLf & "Site:  " & ex.TargetSite.ToString & vbCrLf & "Stack Trace:  " & ex.StackTrace.ToString
        oErr.HandleError(ex, False)
      End Try
    Else

DoURL:
      Try

        Dim bVoidCharge As Boolean = False
        oAuth.InpChargeAmount = dAmountCharged
        If bChargeInTestMode Then
          oAuth.InpLoginID = "38vV43Dhg"
          oAuth.InpTransactionKey = "426bkKR4q6X2F97c"
        Else
          oAuth.InpLoginID = "33WmD7pGbd"
          oAuth.InpTransactionKey = "3q3S4N9Lz2d7Nw3v"
        End If
        oAuth.InpTestMode = bSkipMakingCharge
        oAuth.InpInvoiceNumber = Left(sPaymentLocation, 1) & Left(sPaymentOrigin, 1) & "-" & oBook.Booking_ID_PK__Integer & "-" & oHost.HostName__String

        If sPaymentMethod = "Card" Then

          If sCreditType <> "" Then
            If sPrevTransactionID <> "" Then
              oAuth.InpPrevTransactionID = sPrevTransactionID.ToString
              If sCreditType = "Void" Then
                bVoidCharge = True
                oAuth.InpAuthorizationType = "VOID"
              Else
                oAuth.InpAuthorizationType = "CREDIT"
              End If
            End If
          End If

          '        oAuth.InpCardNumber = "4222222222222"
          If sCCMonth.Trim.Length = 1 Then sCCMonth = "0" & sCCMonth
          sCCYear = sCCYear.Substring(sCCYear.Length - 2, 2)

          oAuth.InpCardNumber = sCCNumber
          oAuth.InpExpMonth = sCCMonth
          oAuth.InpExpYear = sCCYear
          oAuth.InpVerficationCode = sCCVerif
          oAuth.InpFirstName = sCCFirstName
          oAuth.InpLastName = sCCLastName
          oAuth.InpAddress = sCCAddress
          oAuth.InpCity = sCCCity
          oAuth.InpState = sCCState
          oAuth.InpZip = sCCZip

        Else
          oAuth.InpBankABANumber = sBankABANumber
          oAuth.InpBankAccountName = sBankAccountName
          If LCase(sBankAccountType) Like "*personal*" Then
            oAuth.InpBankAccountType = "CHECKING"
          ElseIf LCase(sBankAccountType) Like "*business*" Then
            oAuth.InpBankAccountType = "CHECKING"
          ElseIf LCase(sBankAccountType) Like "*savings*" Then
            oAuth.InpBankAccountType = "SAVINGS"
          End If
          oAuth.InpBankAccountNumber = sBankAccountNumber
          oAuth.InpBankName = sBankName
        End If
        oAuth.PaymentMethod = sPaymentMethod

        Dim sErr As String = ""
        If oAuth.CreateURL() Then
          ProcessCharge = oAuth.MakeCharge(bChargeInTestMode, sErr)
          ' If voiding a previous charge failed, simply issue a credit against the transaction
          If ProcessCharge = False And oAuth.InpAuthorizationType = "VOID" Then
            oAuth.InpAuthorizationType = "CREDIT"
            If oAuth.CreateURL() Then
              ProcessCharge = oAuth.MakeCharge(bChargeInTestMode, sErr)
            End If
          End If
          sPostURLUsed = PostURL
        End If

        sCSTAYResult &= CStayChargeResultText & vbCrLf & sErr
        sUserResult = UserChargeResultText


      Catch ex As Exception
        If bPutErrorInResult Then sCSTAYResult = "ERROR:" & vbCrLf & "Message: " & ex.Message.ToString & vbCrLf & "Site:  " & ex.TargetSite.ToString & vbCrLf & "Stack Trace:  " & ex.StackTrace.ToString
        oErr.HandleError(ex)
      End Try

    End If
  End Function



  Public Sub SendProblemCardEmail(ByVal oPayment As Object, ByVal oAuth As CStayAuthorizeNET,
  ByVal oHost As TableHosts, ByVal oGuest As TableGuests, ByVal oProp As TableProperties)

    Try

      Dim Auth_Info As New System.Collections.Specialized.NameValueCollection, Heading As String = ""
      Dim sMsg As String = ""

      Auth_Info.Add("Reservation ID", oPayment.BookingID__Integer.ToString)
      Auth_Info.Add("Guest", oGuest.GuestFirstName__String & " " & oGuest.GuestLastName__String)
      Auth_Info.Add("Phone", oGuest.HomePhone__String & " (H)  " & oGuest.CellPhone__String & "(Cell)")
      Auth_Info.Add("Property", oProp.PropertyName__String)
      Auth_Info.Add("Host", oHost.HostName__String)
      Auth_Info.Add("Amount Charged", oPayment.Amount__Numeric.ToString)
      Auth_Info.Add("Payment Date", oPayment.Date__Date.ToString)
      If oPayment.Type__String = "Card" Then
        Auth_Info.Add("Credit Card Type", oPayment.CCType__String)
        Auth_Info.Add("Credit Card Name", oPayment.CCName__String)
        Auth_Info.Add("Credit Card Number Encrypted", oPayment.CCNumberEncrypted__String)
        Auth_Info.Add("Credit Card Number", oPayment.CCNumber__String)
        Auth_Info.Add("Credit Card ExpMonth", oPayment.CCExpMonth__String)
        Auth_Info.Add("Credit Card ExpYear", oPayment.CCExpYear__String)

        Auth_Info.Add("Credit Card Address", oPayment.CCAddress__String)
        Auth_Info.Add("Credit Card City", oPayment.CCCity__String)
        Auth_Info.Add("Credit Card State", oPayment.CCState__String)
        Auth_Info.Add("Credit Card Zip", oPayment.CCZip__String)
        Auth_Info.Add("Credit Card Verification", oPayment.CCVerification__String)
      Else
        Auth_Info.Add("Bank Name", oPayment.BankName__String)
        Auth_Info.Add("Bank Account Name", oPayment.BankAccountName__String)
        Auth_Info.Add("Bank Account Number", oPayment.BankAccountNumber__String)
        Auth_Info.Add("Bank Account Type", oPayment.BankAccountType__String)
        Auth_Info.Add("Bank ABA Number", oPayment.BankRoutingNumber__String)
      End If

      Auth_Info.Add("AuthNet Approval Code", oAuth.RespApprovalCode)
      Auth_Info.Add("AuthNet Reason Code", oAuth.RespReasonCode)
      Auth_Info.Add("AuthNet Reason Text", oAuth.RespReasonText)
      Auth_Info.Add("AuthNet Response Code", oAuth.RespCode)
      Auth_Info.Add("AuthNet Trans ID", oAuth.RespTransID)
      Auth_Info.Add("AuthNet AVS Code", oAuth.RespAVSCode)
      Auth_Info.Add("AuthNet AVS Text", oAuth.RespAVSText)
      Auth_Info.Add("AuthNet Verification Code", oAuth.RespCVCCCode)
      Auth_Info.Add("AuthNet Verification Text", oAuth.RespCVCCText)

      sMsg = "<TABLE BORDER=""0"" WIDTH=""100%"" CELLPADDING=""1"" CELLSPACING=""0""><TR><TD bgcolor=""black"" COLSPAN=""2""><FONT face=""Arial"" color=""white""><B> Charge Attempt Information</B></FONT></TD></TR></TABLE>"
      sMsg += oErr.CollectionToHtmlTable(Auth_Info)

      oShared.SendMail(sSystemEmailAddress, sSystemEmailAddress, "Declined/Problem Transaction - " & oPayment.CCName__String, sMsg, "CStay Reservation System", sSystemEmailAddress, True, True)
    Catch ex As Exception
      oErr.LogWebError("Error while doing SendProblemCardEmail()<br><br>" & ex.Message.ToString, ex.Source.ToString, , ex.TargetSite.ToString, , , ex.Source.ToString)
    End Try

  End Sub

  Protected Overrides Sub Finalize()
    oErr = Nothing
    oShared = Nothing
    MyBase.Finalize()
  End Sub

  Public Sub New()
  End Sub

  Public Sub New(ByRef oErrHandler As ErrorHandler, ByVal oSharedUtilities As SharedUtilities)
    oErr = oErrHandler
    oShared = oSharedUtilities
    sSystemEmailAddress = oShared.sSystemEmailAddress
    sConnString = oShared.sConnString
  End Sub

  'Public Sub RemoveCIMRecord(sCIMID As String, bClearAllInGuestRecords As Boolean)

  '  Dim oCIM As New ITDevWorks.CimLib.CimLib(CimLibKey, CimLibKeyName, CimLibKeyEmail), oCCust As ITDevWorks.CimLib.CustomerProfile = Nothing, oCPay As ITDevWorks.CimLib.PaymentProfile = Nothing
  '  oCIM.ApiLogin = "33WmD7pGbd"
  '  oCIM.TransactionKey = "3q3S4N9Lz2d7Nw3v"
  '  oCIM.GatewayUrl = ITDevWorks.CimLib.StandardGatewayUrl.Live

  '  Try

  '    If bClearAllInGuestRecords Then
  '      Dim oHouse As HouseRentalDataContext = MyData.GetHouseRentalContext(True)
  '      Dim oGuests = (From x In oHouse.Guests Where x.AuthNETCustomerProfileID IsNot Nothing), iCount As Integer = 0
  '      If oGuests IsNot Nothing And oGuests.Count > 0 Then
  '        For Each oGuest In oGuests

  '          oCCust = oCIM.GetCustomerProfile(oGuest.AuthNETCustomerProfileID)
  '          For iCount = 0 To oCCust.PaymentProfiles.Count - 1
  '            oCPay = oCCust.PaymentProfiles(iCount)
  '            oCCust.PaymentProfiles.Remove(oCPay)
  '          Next
  '          oCIM.CustomerProfiles.Remove(oCCust)

  '        Next
  '      End If

  '    Else
  '      oCCust = oCIM.GetCustomerProfile(sCIMID)
  '      oCIM.CustomerProfiles.Remove(oCCust)
  '    End If
  '  Catch ex As Exception
  '    oErr.HandleError(ex)
  '  End Try

  'End Sub


End Class
