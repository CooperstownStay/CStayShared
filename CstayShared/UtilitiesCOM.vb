<ComClass(UtilitiesCOM.ClassId, UtilitiesCOM.InterfaceId, UtilitiesCOM.EventsId)> _
Public Class UtilitiesCOM

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
  Public Const ClassId As String = "fa536813-fdd3-4208-a8f0-69a8928f4ee1"
  Public Const InterfaceId As String = "4bedd3d9-e519-4989-9860-58cdb48c06e9"
  Public Const EventsId As String = "f9c60888-e43e-4854-8461-3fa34870f53d"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

  Private sConnString As String = ""
  Private sSMTPServer As String = ""
  Private sSMTPUser As String = ""
  Private sSMTPPass As String = ""
  Private sSystemEmailAddress As String = ""
  Private sErrorEmailAddress As String = ""
  Private sAppEnvironment As String = ""
  Private sAppType As String = ""
  Private sMainWebsite As String = ""
  Private sSecureWebsite As String = ""
  Private bLiveTesting As String = ""
  Private sDateForTesting As String = ""

  Private oUtil As SharedUtilities


  Public Function TestParams(ByRef sResult As String, ByVal sIn As String) As Boolean

    TestParams = LoadCommonParams(sResult, sIn)
    sResult = TestParams.ToString & vbCrLf
    sResult &= "-" & sConnString & vbCrLf
    sResult &= "-" & sSMTPServer & vbCrLf
    sResult &= "-" & sSMTPUser & vbCrLf
    sResult &= "-" & sSMTPPass & vbCrLf
    sResult &= "-" & sSystemEmailAddress & vbCrLf
    sResult &= "-" & sErrorEmailAddress & vbCrLf
    sResult &= "-" & sAppEnvironment & vbCrLf
    sResult &= "-" & sAppType & vbCrLf
    sResult &= "-" & sMainWebsite & vbCrLf
    sResult &= "-" & sSecureWebsite & vbCrLf
    sResult &= "-" & bLiveTesting & vbCrLf


  End Function

  Private Function LoadCommonParams(ByRef sResult As String, ByVal sIn As String) As Boolean
    LoadCommonParams = False
    Try

      If oUtil Is Nothing Then

        Dim sData() As String
        sData = Split(sIn, "|")
        If UBound(sData) > 10 Then

          sConnString = sData(0)
          sSMTPServer = sData(1)
          sSMTPUser = sData(2)
          sSMTPPass = sData(3)
          sSystemEmailAddress = sData(4)
          sErrorEmailAddress = sData(5)
          sAppEnvironment = sData(6)
          sAppType = sData(7)
          sMainWebsite = sData(8)
          sSecureWebsite = sData(9)
          bLiveTesting = sData(10)
          sDateForTesting = sData(11)
        End If
        oUtil = New SharedUtilities(sConnString, sSMTPServer, sSMTPUser, sSMTPPass, sSystemEmailAddress, sErrorEmailAddress, sAppType, sAppEnvironment, sMainWebsite, sSecureWebsite, bLiveTesting, sDateForTesting, "")
        oUtil.oErr.bRunningAsDLL = True
      End If
      LoadCommonParams = True

    Catch ex As Exception
      Dim oErr As New ErrorHandler
      sResult = oErr.GetError(ex)
      oErr = Nothing
    End Try

  End Function



  Public Function SendMail(ByRef sResult As String, ByVal sCommonParams As String, _
  ByVal sSMTPServer As String, _
  ByVal sSMTPUser As String, _
  ByVal sSMTPPassword As String, _
  ByVal sFrom As String, _
  ByVal sFromName As String, _
  ByVal sTo As String, _
  ByVal sToName As String, _
  ByVal sSubject As String, _
  ByVal sBody As String, _
  ByVal BodyFormatIsHTML As Boolean, _
  ByVal bQueueEmail As Boolean, _
  ByVal sSourceTable As String, _
  ByVal iSourceID As Integer, ByVal iSMTPPort As Integer) As Boolean
    Try

      If LoadCommonParams(sResult, sCommonParams) Then
        SendMail = oUtil.SendMailVB6(sResult, sSMTPServer, sSMTPUser, sSMTPPassword, sFrom, sFromName, sTo, sToName, sSubject, sBody, BodyFormatIsHTML, bQueueEmail, sSourceTable, iSourceID, iSMTPPort)
      End If
    Catch ex As Exception
      sResult = oUtil.oErr.GetError(ex)
    End Try

  End Function


  Public Function CreateEmailHTML(ByRef sResult As String, ByVal sCommonParams As String, ByRef sHTMLEmail As String, ByVal sType As String, ByVal sPaymentType As String, ByVal sEmailType As String, ByVal sBookingID As String, ByVal sPaymentID As String, _
   ByVal bClipReservation As String) As Boolean

    CreateEmailHTML = False

    Try

      If LoadCommonParams(sResult, sCommonParams) Then
        If oUtil.CreateEmailHTMLVB6(sResult, sHTMLEmail, sType, sPaymentType, sEmailType, sBookingID, sPaymentID, 0) Then
          If bClipReservation = "True" Then
            sHTMLEmail = oUtil.ClipHTMLSectionToUse(sHTMLEmail, "<MAIN_BODY>")
            sHTMLEmail = oUtil.ClipHTMLSectionToUse(sHTMLEmail, "<main_body>")
          End If
          CreateEmailHTML = True
        End If
      End If

    Catch ex As Exception
      sResult = oUtil.oErr.GetError(ex)
    End Try


  End Function

  Public Function CreateReservationInfoNoCheckIn(ByRef sResult As String, ByVal sCommonParams As String, ByVal sBookingID As String) As Boolean

    CreateReservationInfoNoCheckIn = False
    Try
      If LoadCommonParams(sResult, sCommonParams) Then
        CreateReservationInfoNoCheckIn = oUtil.CreateReservationInfoVB6(sResult, sBookingID, "False")
      End If
    Catch ex As Exception
      sResult = "ERROR:" & vbCrLf & "Message: " & ex.Message.ToString & vbCrLf & "Site:  " & ex.TargetSite.ToString & vbCrLf & "Stack Trace:  " & ex.StackTrace.ToString
      '      sResult = oUtil.oErr.GetError(ex)
    End Try

  End Function
  Public Function CreateReservationInfo(ByRef sResult As String, ByVal sCommonParams As String, ByVal sBookingID As String) As Boolean

    CreateReservationInfo = False
    Try
      If LoadCommonParams(sResult, sCommonParams) Then
        CreateReservationInfo = oUtil.CreateReservationInfoVB6(sResult, sBookingID, "False")
      End If
    Catch ex As Exception
      sResult = "ERROR:" & vbCrLf & "Message: " & ex.Message.ToString & vbCrLf & "Site:  " & ex.TargetSite.ToString & vbCrLf & "Stack Trace:  " & ex.StackTrace.ToString
      '      sResult = oUtil.oErr.GetError(ex)
    End Try

  End Function

  Public Function CreateReservationSummary(ByRef sResult As String, ByVal sCommonParams As String, ByVal sBookingID As String, ByVal bCreateRemainingAmountMessage As String) As Boolean

    CreateReservationSummary = False
    Try
      If LoadCommonParams(sResult, sCommonParams) Then
        CreateReservationSummary = oUtil.CreateReservationSummaryVB6(sResult, sBookingID, bCreateRemainingAmountMessage)
      End If
    Catch ex As Exception
      sResult = oUtil.oErr.GetError(ex)
    End Try

  End Function

  Public Function CreateCancellationPolicy(ByRef sResult As String, ByVal sCommonParams As String, ByVal sRentalFee As String) As Boolean

    CreateCancellationPolicy = False
    Try
      If LoadCommonParams(sResult, sCommonParams) Then
        CreateCancellationPolicy = oUtil.CreateCancellationPolicy(oUtil.CNullD(sRentalFee))
      End If
      CreateCancellationPolicy = True
    Catch ex As Exception
      sResult = oUtil.oErr.GetError(ex)
    End Try
  End Function

  'Public Function CreateCIMPaymentInformation(ByRef sResult As String, ByVal sCommonParams As String, ByVal sGuestID As String, ByVal sGuestName As String, _
  'ByVal sGuestEmail As String, ByRef sCustomerCIMID As String, ByRef sPaymentCIMIDs As String, ByRef sPaymentID As String, ByRef sPaymentCategory As String, _
  'ByRef sCCNumber As String, ByRef sCCMonth As String, ByRef sCCYear As String, ByRef sCCFirstName As String, ByRef sCCLastName As String, _
  'ByRef sCCAddress As String, ByRef sCCCity As String, ByRef sCCState As String, ByRef sCCZip As String, _
  'Optional ByVal bTestMode As Boolean = False) As Boolean

  '  CreateCIMPaymentInformation = False
  '  Try
  '    If LoadCommonParams(sResult, sCommonParams) Then

  '      Dim sData() As String = Nothing, iCount As Integer = 0
  '      Dim sAPaymentID() As String = Nothing, sAPaymentCIMID() As String = Nothing, sACCNumber() As String = Nothing, sACCFirstName() As String = Nothing
  '      Dim sACCLastName() As String = Nothing, sACCAddress() As String = Nothing, sACCCity() As String = Nothing
  '      Dim sACCState() As String = Nothing, sACCZip() As String = Nothing, iIndex As Integer = 0, sAPaymentCategory() As String = Nothing
  '      Dim sACCMonth() As String = Nothing, sACCYear() As String = Nothing

  '      sAPaymentID = Split(sPaymentID, "|")
  '      sAPaymentCategory = Split(sPaymentCategory, "|")
  '      sACCNumber = Split(sCCNumber, "|")
  '      sACCMonth = Split(sCCMonth, "|")
  '      sACCYear = Split(sCCYear, "|")
  '      sACCFirstName = Split(sCCFirstName, "|")
  '      sACCLastName = Split(sCCLastName, "|")
  '      sACCAddress = Split(sCCAddress, "|")
  '      sACCCity = Split(sCCCity, "|")
  '      sACCState = Split(sCCState, "|")
  '      sACCZip = Split(sCCZip, "|")

  '      CreateCIMPaymentInformation = oUtil.CreateCIMPaymentInformationVB6(sResult, sGuestID, sGuestName, _
  '      sGuestEmail, sCustomerCIMID, sAPaymentCIMID, sAPaymentID, sAPaymentCategory, sACCNumber, sACCMonth, sACCYear, sACCFirstName, _
  '      sACCLastName, sACCAddress, sACCCity, sACCState, sACCZip, bTestMode)

  '      sResult = sAPaymentCIMID(0)
  '      If CreateCIMPaymentInformation Then
  '        For iIndex = 0 To UBound(sAPaymentCIMID)
  '          sPaymentCIMIDs = sPaymentCIMIDs & sAPaymentCIMID(iIndex) & "|"
  '        Next
  '        'If sPaymentCIMIDs.Length > 1 Then
  '        '	sPaymentCIMIDs = Left(sPaymentCIMIDs, sPaymentCIMIDs.Length - 1)
  '        'End If
  '      End If

  '    End If
  '  Catch ex As Exception
  '    sResult = oUtil.oErr.GetError(ex)
  '  End Try

  'End Function


  'Public Function GetCIMPaymentInformation(ByRef sResult As String, ByVal sCommonParams As String, ByRef sCustomerCIMID As String, ByRef sPaymentCIMID As String, ByRef sCCPaymentID As String, ByRef sCCNumber As String, _
  'ByRef sCCFirstName As String, ByRef sCCLastName As String, _
  'ByRef sCCAddress As String, ByRef sCCCity As String, ByRef sCCState As String, ByRef sCCZip As String, _
  'ByVal bTestMode As Boolean, ByVal bGetAllPaymentsIfRequestedPaymentMissing As Boolean) As Boolean

  '  GetCIMPaymentInformation = False
  '  Dim sACCPaymentID() As String = Nothing, sACCNumber() As String = Nothing, sACCFirstName() As String = Nothing
  '  Dim sACCLastName() As String = Nothing, sACCAddress() As String = Nothing, sACCCity() As String = Nothing
  '  Dim sACCState() As String = Nothing, sACCZip() As String = Nothing, iIndex As Integer = 0
  '  Try

  '    If LoadCommonParams(sResult, sCommonParams) Then

  '      GetCIMPaymentInformation = oUtil.GetCIMPaymentInformationVB6(sResult, sCustomerCIMID, sPaymentCIMID, sACCPaymentID, sACCNumber, _
  '      sACCFirstName, sACCLastName, sACCAddress, sACCCity, sACCState, sACCZip, bTestMode, bGetAllPaymentsIfRequestedPaymentMissing)

  '    End If

  '    If GetCIMPaymentInformation Then
  '      For iIndex = 0 To UBound(sACCPaymentID)
  '        sCCPaymentID = sCCPaymentID & sACCPaymentID(iIndex) & "|"
  '        sCCNumber = sCCNumber & sACCNumber(iIndex) & "|"
  '        sCCFirstName = sCCFirstName & sACCFirstName(iIndex) & "|"
  '        sCCLastName = sCCLastName & sACCLastName(iIndex) & "|"
  '        sCCAddress = sCCAddress & sACCAddress(iIndex) & "|"
  '        sCCCity = sCCCity & sACCCity(iIndex) & "|"
  '        sCCState = sCCState & sACCState(iIndex) & "|"
  '        sCCZip = sCCZip & sACCZip(iIndex) & "|"
  '      Next
  '      If sCCPaymentID.Length > 1 Then
  '        sCCPaymentID = Left(sCCPaymentID, sCCPaymentID.Length - 1)
  '        sCCNumber = Left(sCCNumber, sCCNumber.Length - 1)
  '        sCCFirstName = Left(sCCFirstName, sCCFirstName.Length - 1)
  '        sCCLastName = Left(sCCLastName, sCCLastName.Length - 1)
  '        sCCAddress = Left(sCCAddress, sCCAddress.Length - 1)
  '        sCCCity = Left(sCCCity, sCCCity.Length - 1)
  '        sCCState = Left(sCCState, sCCState.Length - 1)
  '        sCCZip = Left(sCCZip, sCCZip.Length - 1)
  '      End If
  '    End If

  '  Catch ex As Exception
  '    sResult = oUtil.oErr.GetError(ex)
  '  End Try


  'End Function



  Public Function SavePayment(ByRef sResult As String, ByVal sCommonParams As String, ByRef sAuthNETReason As String, ByRef sAuthNET_AVSText As String, _
   ByRef sCustomerCIMID As String, ByRef sPaymentCIMID As String, ByVal iBookingID As Integer, ByRef iPaymentID As Integer, ByVal sPaymentOrigin As String, _
   ByVal sPaymentLocation As String, ByVal sEmailType As String, ByVal bChargeInTestMode As Boolean, _
   ByVal dPaymentAmount As Double, ByVal dChargeAmount As Double, ByVal sPaymentCategory As String, ByVal sPaymentMethod As String, _
   ByVal sPaymentEmail As String, ByRef iContactLogID As Integer, _
   ByVal sEncryptPassword As String, ByVal sCCType As String, ByVal sCCNumber As String, ByVal sCCFirstName As String, ByVal sCCLastName As String, _
   ByVal sCCExpMonth As String, ByVal sCCExpYear As String, ByVal sCCVerification As String, ByVal sCCAddress As String, ByVal sCCCity As String, _
   ByVal sCCState As String, ByVal sCCZip As String, ByVal sBankName As String, ByVal sBankAccountName As String, ByVal sBankAccountNumber As String, _
   ByVal sBankABANumber As String, ByVal sBankAccountType As String, ByVal bSkipMakingCharge As Boolean, ByVal bPutErrorInResult As Boolean, ByVal sCreditType As String, ByVal sPrevTransactionID As String, _
   ByVal iPrevPaymentID As Integer, ByVal sPrevPaymentCategory As String, ByRef sAuthorizationTypeUsed As String, ByVal sUserName As String, ByVal bUseCIMGateway As Boolean, ByRef bCIMTransactionDeclined As Boolean) As Boolean


    '	ByVal sBankABANumber As String, ByVal sBankAccountType As String, ByVal bSkipMakingCharge As Boolean, _
    'ByVal bIssueCredit As Boolean, ByVal sPrevTransactionID As String, ByVal iPrevPaymentID As Integer, _
    'ByRef sAuthorizationTypeUsed As String, ByVal sUserName As String) As Boolean

    SavePayment = False
        Try


      If LoadCommonParams(sResult, sCommonParams) Then

                SavePayment = oUtil.SavePaymentVB6(sResult, sAuthNETReason, sAuthNET_AVSText, sCustomerCIMID, sPaymentCIMID, iBookingID, iPaymentID, sPaymentOrigin,
                sPaymentLocation, sEmailType, bChargeInTestMode,
                dPaymentAmount, dChargeAmount, sPaymentCategory, sPaymentMethod,
                sPaymentEmail, iContactLogID, False, False,
                sEncryptPassword, sCCType, sCCNumber, sCCFirstName, sCCLastName,
                sCCExpMonth, sCCExpYear, sCCVerification, sCCAddress, sCCCity,
                sCCState, sCCZip, sBankName, sBankAccountName, sBankAccountNumber, sBankABANumber, sBankAccountType, bSkipMakingCharge, bPutErrorInResult,
                sCreditType, sPrevTransactionID, iPrevPaymentID, sPrevPaymentCategory, sAuthorizationTypeUsed, sUserName, bUseCIMGateway, bCIMTransactionDeclined)
            End If

        Catch ex As Exception
            sResult = oUtil.oErr.GetError(ex)
    End Try


  End Function





  Public Function StripHTML(ByVal Source As String) As String
    StripHTML = ""
    Try
      Dim result As String = ""

      '' Remove HTML Development formatting
      'result = Source.Replace(vbCr, " ")                     ' Replace line breaks with space because browsers inserts space
      'result = result.Replace(vbLf, " ")                     ' Replace line breaks with space because browsers inserts space
      'result = result.Replace(vbTab, String.Empty)      ' Remove step-formatting
      'result = System.Text.RegularExpressions.Regex.Replace(result, "( )+", " ")      ' Remove repeating speces becuase browsers ignore them


      ' Remove the header (prepare first by clearing attributes)
      result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*head([^>])*>", "<head>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "(<( )*(/)( )*head( )*>)", "</head>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "(<head>).*(</head>)", String.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

      ' remove all scripts (prepare first by clearing attributes)
      result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*script([^>])*>", "<script>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "(<( )*(/)( )*script( )*>)", "</script>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "(<script>).*(</script>)", String.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

      ' remove all styles (prepare first by clearing attributes)
      result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*style([^>])*>", "<style>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "(<( )*(/)( )*style( )*>)", "</style>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "(<style>).*(</style>)", String.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

      ' insert tabs in spaces of <td> tags
      result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*td([^>])*>", vbTab, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

      ' insert line breaks in places of <BR> and <LI> tags
      result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*br( )*>", vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*li( )*>", vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

      ' insert line paragraphs (double line breaks) in place
      ' if <P>, <DIV> and <TR> tags
      result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*div([^>])*>", vbCr & vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*tr([^>])*>", vbCr & vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*p([^>])*>", vbCr & vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

      ' Remove remaining tags like <a>, links, images,
      ' comments etc - anything that's enclosed inside < >
      result = System.Text.RegularExpressions.Regex.Replace(result, "<[^>]*>", String.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

      ' replace special characters:
      result = System.Text.RegularExpressions.Regex.Replace(result, " ", " ", System.Text.RegularExpressions.RegexOptions.IgnoreCase)

      result = System.Text.RegularExpressions.Regex.Replace(result, "&bull;", " * ", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "&lsaquo;", "<", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "&rsaquo;", ">", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "&trade;", "(tm)", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "&frasl;", "/", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "&lt;", "<", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "&gt;", ">", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "&copy;", "(c)", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "&reg;", "(r)", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      ' Remove all others. More can be added, see
      result = System.Text.RegularExpressions.Regex.Replace(result, "&(.{2,6});", String.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

      ' for testing
      'System.Text.RegularExpressions.Regex.Replace(result,
      '       this.txtRegex.Text,string.Empty,
      '       System.Text.RegularExpressions.RegexOptions.IgnoreCase)

      ' make line breaking consistent
      result = result.Replace(vbLf, vbCr)

      ' Remove extra line breaks and tabs:
      ' replace over 2 breaks with 2 and over 4 tabs with 4.
      ' Prepare first to remove any whitespaces in between
      ' the escaped characters and remove redundant tabs in between line breaks
      result = System.Text.RegularExpressions.Regex.Replace(result, "(\r)( )+(\r)", vbCr & vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "(\t)( )+(\t)", vbTab & vbTab, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "(\t)( )+(\r)", vbTab & vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      result = System.Text.RegularExpressions.Regex.Replace(result, "(\r)( )+(\t)", vbCr & vbTab, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      ' Remove redundant tabs
      result = System.Text.RegularExpressions.Regex.Replace(result, "(\r)(\t)+(\r)", vbCr & vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      ' Remove multiple tabs following a line break with just one tab
      result = System.Text.RegularExpressions.Regex.Replace(result, "(\r)(\t)+", vbCr & vbTab, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      ' Initial replacement target string for line breaks
      Dim breaks As String = vbCr & vbCr & vbCr         ' Initial replacement target string for linebreaks
      Dim tabs As String = vbTab & vbTab & vbTab & vbTab & vbTab             ' Initial replacement target string for tabs

      Dim int As Integer
      For int = 0 To result.Length
        result = result.Replace(breaks, vbCr & vbCr)
        result = result.Replace(tabs, vbTab & vbTab & vbTab & vbTab)
        breaks = breaks + vbCr
        tabs = tabs + vbTab
      Next


      ' Thats it.
      StripHTML = result
    Catch ex As Exception
      StripHTML = "ERROR:" & vbCrLf & "Message: " & ex.Message.ToString & vbCrLf & "Site:  " & ex.TargetSite.ToString & vbCrLf & "Stack Trace:  " & ex.StackTrace.ToString

    End Try
  End Function


  Public Function AES_Encrypt(sInputValue As String, sKey As String) As String
    AES_Encrypt = ""
    Try
      Dim oAES As New MyAES
      AES_Encrypt = oAES.AES_Encrypt(sInputValue, sKey)
    Catch ex As Exception

    End Try
  End Function

  Public Function AES_Decrypt(sInputValue As String, sKey As String) As String
    AES_Decrypt = ""
    Try
      Dim oAES As New MyAES
      AES_Decrypt = oAES.AES_Decrypt(sInputValue, sKey)
    Catch ex As Exception

    End Try
  End Function


  Public Function AuthNET_ProcessCharge() As String

    AuthNET_ProcessCharge = ""
    Try

    Catch ex As Exception
      AuthNET_ProcessCharge = ""

    End Try

  End Function



  Protected Overrides Sub Finalize()
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class


