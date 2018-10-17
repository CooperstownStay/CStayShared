Imports CstayShared

Public Class Form1

  Private Sub btnDoIt_Click(sender As System.Object, e As System.EventArgs) Handles btnDoIt.Click

    Dim oTest As New CstayShared.UtilitiesCOM
    'Dim sInp As String = "Test Value", sEnc As String = "", sOut As String = ""
    'sEnc = oTest.AES_Encrypt(sInp, "freed0m")
    'sOut = oTest.AES_Decrypt(sEnc, "freed0m")
    Dim bResult As Boolean = False

    Dim sRequest As String = "x_login=38vV43Dhg&x_tran_key=426bkKR4q6X2F97c&x_version=3.1&x_delim_data=TRUE&x_delim_char=|&x_relay_response=False&x_amount=10&x_invoice_num=OR-76158-Schmitt%2c+Lester&x_first_name=Byron&x_last_name=Richard&x_address=288+Deerfield+Estate+Road&x_city=Boone&x_state=NC&x_zip=28607&x_method=CC&x_card_num=4147202093358305&x_exp_date=1118&x_type=AUTH_CAPTURE"
    Dim oAuth As New CStayAuthorizeNET()
    'oAuth.RetrieveHtmlPage(sRequest, True,)


    Dim sResult As String = ""
    Dim sCStaySharedParams As String = "Data Source=cstay-server.cooperstownstay.com,51433;Initial Catalog=HouseRentalTest;Persist Security Info=True;User ID=sa;Password=4Gsltw316;Connect Timeout=60;|smtp.stny.rr.com|||find@cooperstownstay.com|developer@cooperstownstay.com|||http://www.CooperstownStay.com|https://www.CStayReserve.com|False||False|False"
    Dim sAuthNETResult As String = ""
    Dim sAuthNET_AVSText As String = ""
    Dim sCustomerCIMID As String = ""
    Dim sPaymentCIMID As String = ""
        Dim iBookingID As Integer = 81883
        Dim iPaymentID As Integer = 0
    Dim sPaymentCategory As String = "Deposit"
    Dim bChargeInTestMode As Boolean = False
    Dim dPaymentAmount As Double = 100
    Dim dChargeAmount As Double = 100
    Dim sPaymentMethod As String = "Card"
    Dim sPaymentEmail As String = "developer@cooperstownstay.com"
    Dim iContactLogID As Integer = 0
    Dim sEncryptPassword As String = "freed0m"
    Dim sCCType As String = "VISA"
    Dim sCCNumber As String = "4147202093358305"
    Dim sCCFirstName As String = "Byron"
    Dim sCCLastName As String = "Richard"
    Dim sCCExpMonth As String = "11"
    Dim sCCExpYear As String = "21"
    Dim sCCVerification As String = ""
    Dim sCCAddress As String = "288 Deerfield Estates Road"
    Dim sCCCity As String = "Boone"
    Dim sCCState As String = "NC"
    Dim sCCZip As String = "28607"
    Dim sBankName As String = ""
    Dim sBankAccountName As String = ""
    Dim sBankAccountNumber As String = ""
    Dim sBankABANumber As String = ""
    Dim sBankAccountType As String = ""
    Dim bSkipMakingCharge As Boolean = False
    Dim bPutErrorInResult As Boolean = True
    Dim sCreditType As String = ""
    Dim sPrevTransactionID As String = ""
    Dim iPrevPaymentID As Integer = 0
    Dim sPrevPaymentCategory As String = "Deposit"
    Dim sAuthorizationTypeUsed As String = ""
    Dim sUserName As String = "br"
    Dim bMyUseCIMGateway As Boolean = True
    Dim bCIMTransactionDeclined As Boolean = False

        ' For Voiding a charge
        'dPaymentAmount = 100
        'dChargeAmount = 100
        'sPaymentCategory = "Refund"
        'sCreditType = "Void"
        'sCustomerCIMID = "1935697388"
        'sPaymentCIMID = "1949307153"
        'iPrevPaymentID = 57270
        'sPrevTransactionID = "61305351677"


        bResult = oTest.SavePayment(sResult, sCStaySharedParams, sAuthNETResult, sAuthNET_AVSText, sCustomerCIMID, sPaymentCIMID,
            iBookingID, iPaymentID, sPaymentCategory, "Phone", "Reservation Payment Email", bChargeInTestMode,
            dPaymentAmount, dChargeAmount, sPaymentCategory, sPaymentMethod,
            sPaymentEmail, iContactLogID,
            sEncryptPassword, sCCType, sCCNumber, sCCFirstName, sCCLastName,
            sCCExpMonth, sCCExpYear, sCCVerification, sCCAddress, sCCCity,
            sCCState, sCCZip, sBankName, sBankAccountName, sBankAccountNumber,
            sBankABANumber, sBankAccountType, bSkipMakingCharge, bPutErrorInResult,
            sCreditType, sPrevTransactionID, iPrevPaymentID, sPrevPaymentCategory,
            sAuthorizationTypeUsed, sUserName, bMyUseCIMGateway, bCIMTransactionDeclined)


  End Sub
End Class
