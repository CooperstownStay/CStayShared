Public Class TableBookings

  Public Connection As New System.Data.SqlClient.SqlConnection()
Public Transaction As System.Data.SqlClient.SqlTransaction
Public SelectedData As Object
Public CurrentRow As  Object
Public ConnectionString as String=""
Public CurrentRecordNumber As Integer = 0
Public oUtil As DBUtilities
Public Sub New(Optional ByVal bBeginTransaction as Boolean=False)

oUtil = New DBUtilities
ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iBooking_ID As Int32
  Private sInsUpdBooking_ID As String
  Property Booking_ID_PK__Integer() As Int32
    Get
      Return iBooking_ID
    End Get
    Set(ByVal Value As Int32)
      iBooking_ID = Value
      sInsUpdBooking_ID = oUtil.FixParam(iBooking_ID, True)
    End Set
  End Property

  Private iProperty_ID As Int32
  Private sInsUpdProperty_ID As String
  Property Property_ID__Integer() As Int32
    Get
      Return iProperty_ID
    End Get
    Set(ByVal Value As Int32)
      iProperty_ID = Value
      sInsUpdProperty_ID = oUtil.FixParam(iProperty_ID, True)
    End Set
  End Property

  Private iHost_ID As Int32
  Private sInsUpdHost_ID As String
  Property Host_ID__Integer() As Int32
    Get
      Return iHost_ID
    End Get
    Set(ByVal Value As Int32)
      iHost_ID = Value
      sInsUpdHost_ID = oUtil.FixParam(iHost_ID, True)
    End Set
  End Property

  Private iGuest_ID As Int32
  Private sInsUpdGuest_ID As String
  Property Guest_ID__Integer() As Int32
    Get
      Return iGuest_ID
    End Get
    Set(ByVal Value As Int32)
      iGuest_ID = Value
      sInsUpdGuest_ID = oUtil.FixParam(iGuest_ID, True)
    End Set
  End Property

  Private sArriveDate As String
  Private sInsUpdArriveDate As String
  Property ArriveDate__Date() As String
    Get
      Return sArriveDate
    End Get
    Set(ByVal Value As String)
      sArriveDate = Value
      sInsUpdArriveDate = oUtil.FixParam(sArriveDate, True)
    End Set
  End Property

  Private sDepartDate As String
  Private sInsUpdDepartDate As String
  Property DepartDate__Date() As String
    Get
      Return sDepartDate
    End Get
    Set(ByVal Value As String)
      sDepartDate = Value
      sInsUpdDepartDate = oUtil.FixParam(sDepartDate, True)
    End Set
  End Property

  Private sStatus As String
  Private sInsUpdStatus As String
  Property Status__String() As String
    Get
      Return sStatus
    End Get
    Set(ByVal Value As String)
      sStatus = Value
      sInsUpdStatus = oUtil.FixParam(sStatus, True)
    End Set
  End Property

  Private sHostTelephone As String
  Private sInsUpdHostTelephone As String
  Property HostTelephone__String() As String
    Get
      Return sHostTelephone
    End Get
    Set(ByVal Value As String)
      sHostTelephone = Value
      sInsUpdHostTelephone = oUtil.FixParam(sHostTelephone, True)
    End Set
  End Property

  Private dRate As Double
  Private sInsUpdRate As String
  Property Rate__Numeric() As Double
    Get
      Return dRate
    End Get
    Set(ByVal Value As Double)
      dRate = Value
      sInsUpdRate = oUtil.FixParam(dRate, True)
    End Set
  End Property

  Private sRequestReceived As String
  Private sInsUpdRequestReceived As String
  Property RequestReceived__Date() As String
    Get
      Return sRequestReceived
    End Get
    Set(ByVal Value As String)
      sRequestReceived = Value
      sInsUpdRequestReceived = oUtil.FixParam(sRequestReceived, True)
    End Set
  End Property

  Private dHostPaidAmtDep As Double
  Private sInsUpdHostPaidAmtDep As String
  Property HostPaidAmtDep__Numeric() As Double
    Get
      Return dHostPaidAmtDep
    End Get
    Set(ByVal Value As Double)
      dHostPaidAmtDep = Value
      sInsUpdHostPaidAmtDep = oUtil.FixParam(dHostPaidAmtDep, True)
    End Set
  End Property

  Private dHostPaidAmtDepFromCredit As Double
  Private sInsUpdHostPaidAmtDepFromCredit As String
  Property HostPaidAmtDepFromCredit__Numeric() As Double
    Get
      Return dHostPaidAmtDepFromCredit
    End Get
    Set(ByVal Value As Double)
      dHostPaidAmtDepFromCredit = Value
      sInsUpdHostPaidAmtDepFromCredit = oUtil.FixParam(dHostPaidAmtDepFromCredit, True)
    End Set
  End Property

  Private sDepositDue As String
  Private sInsUpdDepositDue As String
  Property DepositDue__Date() As String
    Get
      Return sDepositDue
    End Get
    Set(ByVal Value As String)
      sDepositDue = Value
      sInsUpdDepositDue = oUtil.FixParam(sDepositDue, True)
    End Set
  End Property

  Private sDepositReceived As String
  Private sInsUpdDepositReceived As String
  Property DepositReceived__Date() As String
    Get
      Return sDepositReceived
    End Get
    Set(ByVal Value As String)
      sDepositReceived = Value
      sInsUpdDepositReceived = oUtil.FixParam(sDepositReceived, True)
    End Set
  End Property

  Private dDepositAmount As Double
  Private sInsUpdDepositAmount As String
  Property DepositAmount__Numeric() As Double
    Get
      Return dDepositAmount
    End Get
    Set(ByVal Value As Double)
      dDepositAmount = Value
      sInsUpdDepositAmount = oUtil.FixParam(dDepositAmount, True)
    End Set
  End Property

  Private sBalanceDueDate As String
  Private sInsUpdBalanceDueDate As String
  Property BalanceDueDate__Date() As String
    Get
      Return sBalanceDueDate
    End Get
    Set(ByVal Value As String)
      sBalanceDueDate = Value
      sInsUpdBalanceDueDate = oUtil.FixParam(sBalanceDueDate, True)
    End Set
  End Property

  Private dBalanceDue As Double
  Private sInsUpdBalanceDue As String
  Property BalanceDue__Numeric() As Double
    Get
      Return dBalanceDue
    End Get
    Set(ByVal Value As Double)
      dBalanceDue = Value
      sInsUpdBalanceDue = oUtil.FixParam(dBalanceDue, True)
    End Set
  End Property

  Private sConfirmationSent As String
  Private sInsUpdConfirmationSent As String
  Property ConfirmationSent__Date() As String
    Get
      Return sConfirmationSent
    End Get
    Set(ByVal Value As String)
      sConfirmationSent = Value
      sInsUpdConfirmationSent = oUtil.FixParam(sConfirmationSent, True)
    End Set
  End Property

  Private sBalanceReceived As String
  Private sInsUpdBalanceReceived As String
  Property BalanceReceived__Date() As String
    Get
      Return sBalanceReceived
    End Get
    Set(ByVal Value As String)
      sBalanceReceived = Value
      sInsUpdBalanceReceived = oUtil.FixParam(sBalanceReceived, True)
    End Set
  End Property

  Private dBalanceAmount As Double
  Private sInsUpdBalanceAmount As String
  Property BalanceAmount__Numeric() As Double
    Get
      Return dBalanceAmount
    End Get
    Set(ByVal Value As Double)
      dBalanceAmount = Value
      sInsUpdBalanceAmount = oUtil.FixParam(dBalanceAmount, True)
    End Set
  End Property

  Private sHostPdDate As String
  Private sInsUpdHostPdDate As String
  Property HostPdDate__Date() As String
    Get
      Return sHostPdDate
    End Get
    Set(ByVal Value As String)
      sHostPdDate = Value
      sInsUpdHostPdDate = oUtil.FixParam(sHostPdDate, True)
    End Set
  End Property

  Private dHostPaidAmtBal As Double
  Private sInsUpdHostPaidAmtBal As String
  Property HostPaidAmtBal__Numeric() As Double
    Get
      Return dHostPaidAmtBal
    End Get
    Set(ByVal Value As Double)
      dHostPaidAmtBal = Value
      sInsUpdHostPaidAmtBal = oUtil.FixParam(dHostPaidAmtBal, True)
    End Set
  End Property

  Private dHostPaidAmtBalFromCredit As Double
  Private sInsUpdHostPaidAmtBalFromCredit As String
  Property HostPaidAmtBalFromCredit__Numeric() As Double
    Get
      Return dHostPaidAmtBalFromCredit
    End Get
    Set(ByVal Value As Double)
      dHostPaidAmtBalFromCredit = Value
      sInsUpdHostPaidAmtBalFromCredit = oUtil.FixParam(dHostPaidAmtBalFromCredit, True)
    End Set
  End Property

  Private dCommission As Double
  Private sInsUpdCommission As String
  Property Commission__Numeric() As Double
    Get
      Return dCommission
    End Get
    Set(ByVal Value As Double)
      dCommission = Value
      sInsUpdCommission = oUtil.FixParam(dCommission, True)
    End Set
  End Property

  Private sBookingNotes As String
  Private sInsUpdBookingNotes As String
  Property BookingNotes__String() As String
    Get
      Return sBookingNotes
    End Get
    Set(ByVal Value As String)
      sBookingNotes = Value
      sInsUpdBookingNotes = oUtil.FixParam(sBookingNotes, True)
    End Set
  End Property

  Private sDepCardType As String
  Private sInsUpdDepCardType As String
  Property DepCardType__String() As String
    Get
      Return sDepCardType
    End Get
    Set(ByVal Value As String)
      sDepCardType = Value
      sInsUpdDepCardType = oUtil.FixParam(sDepCardType, True)
    End Set
  End Property

  Private sDepCardNumber As String
  Private sInsUpdDepCardNumber As String
  Property DepCardNumber__String() As String
    Get
      Return sDepCardNumber
    End Get
    Set(ByVal Value As String)
      sDepCardNumber = Value
      sInsUpdDepCardNumber = oUtil.FixParam(sDepCardNumber, True)
    End Set
  End Property

  Private sDepCardNumberEncrypted As String
  Private sInsUpdDepCardNumberEncrypted As String
  Property DepCardNumberEncrypted__String() As String
    Get
      Return sDepCardNumberEncrypted
    End Get
    Set(ByVal Value As String)
      sDepCardNumberEncrypted = Value
      sInsUpdDepCardNumberEncrypted = oUtil.FixParam(sDepCardNumberEncrypted, True)
    End Set
  End Property

  Private sDepCardConfirm As String
  Private sInsUpdDepCardConfirm As String
  Property DepCardConfirm__String() As String
    Get
      Return sDepCardConfirm
    End Get
    Set(ByVal Value As String)
      sDepCardConfirm = Value
      sInsUpdDepCardConfirm = oUtil.FixParam(sDepCardConfirm, True)
    End Set
  End Property

  Private sBalCardType As String
  Private sInsUpdBalCardType As String
  Property BalCardType__String() As String
    Get
      Return sBalCardType
    End Get
    Set(ByVal Value As String)
      sBalCardType = Value
      sInsUpdBalCardType = oUtil.FixParam(sBalCardType, True)
    End Set
  End Property

  Private sBalCardNumber As String
  Private sInsUpdBalCardNumber As String
  Property BalCardNumber__String() As String
    Get
      Return sBalCardNumber
    End Get
    Set(ByVal Value As String)
      sBalCardNumber = Value
      sInsUpdBalCardNumber = oUtil.FixParam(sBalCardNumber, True)
    End Set
  End Property

  Private sBalCardNumberEncrypted As String
  Private sInsUpdBalCardNumberEncrypted As String
  Property BalCardNumberEncrypted__String() As String
    Get
      Return sBalCardNumberEncrypted
    End Get
    Set(ByVal Value As String)
      sBalCardNumberEncrypted = Value
      sInsUpdBalCardNumberEncrypted = oUtil.FixParam(sBalCardNumberEncrypted, True)
    End Set
  End Property

  Private sBalCardConfirm As String
  Private sInsUpdBalCardConfirm As String
  Property BalCardConfirm__String() As String
    Get
      Return sBalCardConfirm
    End Get
    Set(ByVal Value As String)
      sBalCardConfirm = Value
      sInsUpdBalCardConfirm = oUtil.FixParam(sBalCardConfirm, True)
    End Set
  End Property

  Private dAutoCharge As Double
  Private sInsUpdAutoCharge As String
  Property AutoCharge__Integer() As Double
    Get
      Return dAutoCharge
    End Get
    Set(ByVal Value As Double)
      dAutoCharge = Value
      sInsUpdAutoCharge = oUtil.FixParam(dAutoCharge, True)
    End Set
  End Property

  Private iAdults As Int32
  Private sInsUpdAdults As String
  Property Adults__Integer() As Int32
    Get
      Return iAdults
    End Get
    Set(ByVal Value As Int32)
      iAdults = Value
      sInsUpdAdults = oUtil.FixParam(iAdults, True)
    End Set
  End Property

  Private iChildren As Int32
  Private sInsUpdChildren As String
  Property Children__Integer() As Int32
    Get
      Return iChildren
    End Get
    Set(ByVal Value As Int32)
      iChildren = Value
      sInsUpdChildren = oUtil.FixParam(iChildren, True)
    End Set
  End Property

  Private iTeens As Int32
  Private sInsUpdTeens As String
  Property Teens__Integer() As Int32
    Get
      Return iTeens
    End Get
    Set(ByVal Value As Int32)
      iTeens = Value
      sInsUpdTeens = oUtil.FixParam(iTeens, True)
    End Set
  End Property

  Private dTaxDue As Double
  Private sInsUpdTaxDue As String
  Property TaxDue__Numeric() As Double
    Get
      Return dTaxDue
    End Get
    Set(ByVal Value As Double)
      dTaxDue = Value
      sInsUpdTaxDue = oUtil.FixParam(dTaxDue, True)
    End Set
  End Property

  Private dRentalBalance As Double
  Private sInsUpdRentalBalance As String
  Property RentalBalance__Numeric() As Double
    Get
      Return dRentalBalance
    End Get
    Set(ByVal Value As Double)
      dRentalBalance = Value
      sInsUpdRentalBalance = oUtil.FixParam(dRentalBalance, True)
    End Set
  End Property

  Private sDepCheckName As String
  Private sInsUpdDepCheckName As String
  Property DepCheckName__String() As String
    Get
      Return sDepCheckName
    End Get
    Set(ByVal Value As String)
      sDepCheckName = Value
      sInsUpdDepCheckName = oUtil.FixParam(sDepCheckName, True)
    End Set
  End Property

  Private sDepCheckNumber As String
  Private sInsUpdDepCheckNumber As String
  Property DepCheckNumber__String() As String
    Get
      Return sDepCheckNumber
    End Get
    Set(ByVal Value As String)
      sDepCheckNumber = Value
      sInsUpdDepCheckNumber = oUtil.FixParam(sDepCheckNumber, True)
    End Set
  End Property

  Private sBalCheckName As String
  Private sInsUpdBalCheckName As String
  Property BalCheckName__String() As String
    Get
      Return sBalCheckName
    End Get
    Set(ByVal Value As String)
      sBalCheckName = Value
      sInsUpdBalCheckName = oUtil.FixParam(sBalCheckName, True)
    End Set
  End Property

  Private sBalCheckNumber As String
  Private sInsUpdBalCheckNumber As String
  Property BalCheckNumber__String() As String
    Get
      Return sBalCheckNumber
    End Get
    Set(ByVal Value As String)
      sBalCheckNumber = Value
      sInsUpdBalCheckNumber = oUtil.FixParam(sBalCheckNumber, True)
    End Set
  End Property

  Private sAddDate As String
  Private sInsUpdAddDate As String
  Property AddDate__Date() As String
    Get
      Return sAddDate
    End Get
    Set(ByVal Value As String)
      sAddDate = Value
      sInsUpdAddDate = oUtil.FixParam(sAddDate, True)
    End Set
  End Property

  Private sLocationBooked As String
  Private sInsUpdLocationBooked As String
  Property LocationBooked__String() As String
    Get
      Return sLocationBooked
    End Get
    Set(ByVal Value As String)
      sLocationBooked = Value
      sInsUpdLocationBooked = oUtil.FixParam(sLocationBooked, True)
    End Set
  End Property

  Private sOrigDepCardNumber As String
  Private sInsUpdOrigDepCardNumber As String
  Property OrigDepCardNumber__String() As String
    Get
      Return sOrigDepCardNumber
    End Get
    Set(ByVal Value As String)
      sOrigDepCardNumber = Value
      sInsUpdOrigDepCardNumber = oUtil.FixParam(sOrigDepCardNumber, True)
    End Set
  End Property

  Private sOrigBalCardNumber As String
  Private sInsUpdOrigBalCardNumber As String
  Property OrigBalCardNumber__String() As String
    Get
      Return sOrigBalCardNumber
    End Get
    Set(ByVal Value As String)
      sOrigBalCardNumber = Value
      sInsUpdOrigBalCardNumber = oUtil.FixParam(sOrigBalCardNumber, True)
    End Set
  End Property

  Private iCancellationHostPaymentProcessed As Int32
  Private sInsUpdCancellationHostPaymentProcessed As String
  Property CancellationHostPaymentProcessed__Integer() As Int32
    Get
      Return iCancellationHostPaymentProcessed
    End Get
    Set(ByVal Value As Int32)
      iCancellationHostPaymentProcessed = Value
      sInsUpdCancellationHostPaymentProcessed = oUtil.FixParam(iCancellationHostPaymentProcessed, True)
    End Set
  End Property

  Private sLateCancellationRefundDate As String
  Private sInsUpdLateCancellationRefundDate As String
  Property LateCancellationRefundDate__Date() As String
    Get
      Return sLateCancellationRefundDate
    End Get
    Set(ByVal Value As String)
      sLateCancellationRefundDate = Value
      sInsUpdLateCancellationRefundDate = oUtil.FixParam(sLateCancellationRefundDate, True)
    End Set
  End Property

  Private sLateCancellationRefundStatus As String
  Private sInsUpdLateCancellationRefundStatus As String
  Property LateCancellationRefundStatus__String() As String
    Get
      Return sLateCancellationRefundStatus
    End Get
    Set(ByVal Value As String)
      sLateCancellationRefundStatus = Value
      sInsUpdLateCancellationRefundStatus = oUtil.FixParam(sLateCancellationRefundStatus, True)
    End Set
  End Property

  Private dLateCancellationRefundAmount As Double
  Private sInsUpdLateCancellationRefundAmount As String
  Property LateCancellationRefundAmount__Numeric() As Double
    Get
      Return dLateCancellationRefundAmount
    End Get
    Set(ByVal Value As Double)
      dLateCancellationRefundAmount = Value
      sInsUpdLateCancellationRefundAmount = oUtil.FixParam(dLateCancellationRefundAmount, True)
    End Set
  End Property

  Private dLateCancellationFeeAmount As Double
  Private sInsUpdLateCancellationFeeAmount As String
  Property LateCancellationFeeAmount__Numeric() As Double
    Get
      Return dLateCancellationFeeAmount
    End Get
    Set(ByVal Value As Double)
      dLateCancellationFeeAmount = Value
      sInsUpdLateCancellationFeeAmount = oUtil.FixParam(dLateCancellationFeeAmount, True)
    End Set
  End Property

  Private iLateCancellationRefundReBookingID As Int32
  Private sInsUpdLateCancellationRefundReBookingID As String
  Property LateCancellationRefundReBookingID__Integer() As Int32
    Get
      Return iLateCancellationRefundReBookingID
    End Get
    Set(ByVal Value As Int32)
      iLateCancellationRefundReBookingID = Value
      sInsUpdLateCancellationRefundReBookingID = oUtil.FixParam(iLateCancellationRefundReBookingID, True)
    End Set
  End Property

  Private sCancellationDate As String
  Private sInsUpdCancellationDate As String
  Property CancellationDate__Date() As String
    Get
      Return sCancellationDate
    End Get
    Set(ByVal Value As String)
      sCancellationDate = Value
      sInsUpdCancellationDate = oUtil.FixParam(sCancellationDate, True)
    End Set
  End Property

  Private iFinalBalancePaymentMadeOnline As Int32
  Private sInsUpdFinalBalancePaymentMadeOnline As String
  Property FinalBalancePaymentMadeOnline__Integer() As Int32
    Get
      Return iFinalBalancePaymentMadeOnline
    End Get
    Set(ByVal Value As Int32)
      iFinalBalancePaymentMadeOnline = Value
      sInsUpdFinalBalancePaymentMadeOnline = oUtil.FixParam(iFinalBalancePaymentMadeOnline, True)
    End Set
  End Property

  Private iFinalDepositPaymentMadeOnline As Int32
  Private sInsUpdFinalDepositPaymentMadeOnline As String
  Property FinalDepositPaymentMadeOnline__Integer() As Int32
    Get
      Return iFinalDepositPaymentMadeOnline
    End Get
    Set(ByVal Value As Int32)
      iFinalDepositPaymentMadeOnline = Value
      sInsUpdFinalDepositPaymentMadeOnline = oUtil.FixParam(iFinalDepositPaymentMadeOnline, True)
    End Set
  End Property

  Private dDepositAmountPaid As Double
  Private sInsUpdDepositAmountPaid As String
  Property DepositAmountPaid__Numeric() As Double
    Get
      Return dDepositAmountPaid
    End Get
    Set(ByVal Value As Double)
      dDepositAmountPaid = Value
      sInsUpdDepositAmountPaid = oUtil.FixParam(dDepositAmountPaid, True)
    End Set
  End Property

  Private sConfirmationEmail As String
  Private sInsUpdConfirmationEmail As String
  Property ConfirmationEmail__String() As String
    Get
      Return sConfirmationEmail
    End Get
    Set(ByVal Value As String)
      sConfirmationEmail = Value
      sInsUpdConfirmationEmail = oUtil.FixParam(sConfirmationEmail, True)
    End Set
  End Property

  Private sConfirmationEmailSentDate As String
  Private sInsUpdConfirmationEmailSentDate As String
  Property ConfirmationEmailSentDate__Date() As String
    Get
      Return sConfirmationEmailSentDate
    End Get
    Set(ByVal Value As String)
      sConfirmationEmailSentDate = Value
      sInsUpdConfirmationEmailSentDate = oUtil.FixParam(sConfirmationEmailSentDate, True)
    End Set
  End Property

  Private iQBActivityIDFirstPayment As Int32
  Private sInsUpdQBActivityIDFirstPayment As String
  Property QBActivityIDFirstPayment__Integer() As Int32
    Get
      Return iQBActivityIDFirstPayment
    End Get
    Set(ByVal Value As Int32)
      iQBActivityIDFirstPayment = Value
      sInsUpdQBActivityIDFirstPayment = oUtil.FixParam(iQBActivityIDFirstPayment, True)
    End Set
  End Property

  Private iQBActivityIDBalancePayment As Int32
  Private sInsUpdQBActivityIDBalancePayment As String
  Property QBActivityIDBalancePayment__Integer() As Int32
    Get
      Return iQBActivityIDBalancePayment
    End Get
    Set(ByVal Value As Int32)
      iQBActivityIDBalancePayment = Value
      sInsUpdQBActivityIDBalancePayment = oUtil.FixParam(iQBActivityIDBalancePayment, True)
    End Set
  End Property

  Private iQBActivityIDFullPayment As Int32
  Private sInsUpdQBActivityIDFullPayment As String
  Property QBActivityIDFullPayment__Integer() As Int32
    Get
      Return iQBActivityIDFullPayment
    End Get
    Set(ByVal Value As Int32)
      iQBActivityIDFullPayment = Value
      sInsUpdQBActivityIDFullPayment = oUtil.FixParam(iQBActivityIDFullPayment, True)
    End Set
  End Property

  Private sPrivateNotes As String
  Private sInsUpdPrivateNotes As String
  Property PrivateNotes__String() As String
    Get
      Return sPrivateNotes
    End Get
    Set(ByVal Value As String)
      sPrivateNotes = Value
      sInsUpdPrivateNotes = oUtil.FixParam(sPrivateNotes, True)
    End Set
  End Property

  Public Sub Clear()
    iBooking_ID = 0
    sInsUpdBooking_ID = ""
    iProperty_ID = 0
    sInsUpdProperty_ID = ""
    iHost_ID = 0
    sInsUpdHost_ID = ""
    iGuest_ID = 0
    sInsUpdGuest_ID = ""
    sArriveDate = ""
    sInsUpdArriveDate = ""
    sDepartDate = ""
    sInsUpdDepartDate = ""
    sStatus = ""
    sInsUpdStatus = ""
    sHostTelephone = ""
    sInsUpdHostTelephone = ""
    dRate = 0.0
    sInsUpdRate = ""
    sRequestReceived = ""
    sInsUpdRequestReceived = ""
    dHostPaidAmtDep = 0.0
    sInsUpdHostPaidAmtDep = ""
    dHostPaidAmtDepFromCredit = 0.0
    sInsUpdHostPaidAmtDepFromCredit = ""
    sDepositDue = ""
    sInsUpdDepositDue = ""
    sDepositReceived = ""
    sInsUpdDepositReceived = ""
    dDepositAmount = 0.0
    sInsUpdDepositAmount = ""
    sBalanceDueDate = ""
    sInsUpdBalanceDueDate = ""
    dBalanceDue = 0.0
    sInsUpdBalanceDue = ""
    sConfirmationSent = ""
    sInsUpdConfirmationSent = ""
    sBalanceReceived = ""
    sInsUpdBalanceReceived = ""
    dBalanceAmount = 0.0
    sInsUpdBalanceAmount = ""
    sHostPdDate = ""
    sInsUpdHostPdDate = ""
    dHostPaidAmtBal = 0.0
    sInsUpdHostPaidAmtBal = ""
    dHostPaidAmtBalFromCredit = 0.0
    sInsUpdHostPaidAmtBalFromCredit = ""
    dCommission = 0.0
    sInsUpdCommission = ""
    sBookingNotes = ""
    sInsUpdBookingNotes = ""
    sDepCardType = ""
    sInsUpdDepCardType = ""
    sDepCardNumber = ""
    sInsUpdDepCardNumber = ""
    sDepCardNumberEncrypted = ""
    sInsUpdDepCardNumberEncrypted = ""
    sDepCardConfirm = ""
    sInsUpdDepCardConfirm = ""
    sBalCardType = ""
    sInsUpdBalCardType = ""
    sBalCardNumber = ""
    sInsUpdBalCardNumber = ""
    sBalCardNumberEncrypted = ""
    sInsUpdBalCardNumberEncrypted = ""
    sBalCardConfirm = ""
    sInsUpdBalCardConfirm = ""
    dAutoCharge = 0.0
    sInsUpdAutoCharge = ""
    iAdults = 0
    sInsUpdAdults = ""
    iChildren = 0
    sInsUpdChildren = ""
    iTeens = 0
    sInsUpdTeens = ""
    dTaxDue = 0.0
    sInsUpdTaxDue = ""
    dRentalBalance = 0.0
    sInsUpdRentalBalance = ""
    sDepCheckName = ""
    sInsUpdDepCheckName = ""
    sDepCheckNumber = ""
    sInsUpdDepCheckNumber = ""
    sBalCheckName = ""
    sInsUpdBalCheckName = ""
    sBalCheckNumber = ""
    sInsUpdBalCheckNumber = ""
    sAddDate = ""
    sInsUpdAddDate = ""
    sLocationBooked = ""
    sInsUpdLocationBooked = ""
    sOrigDepCardNumber = ""
    sInsUpdOrigDepCardNumber = ""
    sOrigBalCardNumber = ""
    sInsUpdOrigBalCardNumber = ""
    iCancellationHostPaymentProcessed = 0
    sInsUpdCancellationHostPaymentProcessed = ""
    sLateCancellationRefundDate = ""
    sInsUpdLateCancellationRefundDate = ""
    sLateCancellationRefundStatus = ""
    sInsUpdLateCancellationRefundStatus = ""
    dLateCancellationRefundAmount = 0.0
    sInsUpdLateCancellationRefundAmount = ""
    dLateCancellationFeeAmount = 0.0
    sInsUpdLateCancellationFeeAmount = ""
    iLateCancellationRefundReBookingID = 0
    sInsUpdLateCancellationRefundReBookingID = ""
    sCancellationDate = ""
    sInsUpdCancellationDate = ""
    iFinalBalancePaymentMadeOnline = 0
    sInsUpdFinalBalancePaymentMadeOnline = ""
    iFinalDepositPaymentMadeOnline = 0
    sInsUpdFinalDepositPaymentMadeOnline = ""
    dDepositAmountPaid = 0.0
    sInsUpdDepositAmountPaid = ""
    sConfirmationEmail = ""
    sInsUpdConfirmationEmail = ""
    sConfirmationEmailSentDate = ""
    sInsUpdConfirmationEmailSentDate = ""
    iQBActivityIDFirstPayment = 0
    sInsUpdQBActivityIDFirstPayment = ""
    iQBActivityIDBalancePayment = 0
    sInsUpdQBActivityIDBalancePayment = ""
    iQBActivityIDFullPayment = 0
    sInsUpdQBActivityIDFullPayment = ""
    sPrivateNotes = ""
    sInsUpdPrivateNotes = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdBooking_ID.ToString = "'-12345'" Then sb.Append("Booking_ID,")
        If sInsUpdProperty_ID.ToString = "'-12345'" Then sb.Append("Property_ID,")
        If sInsUpdHost_ID.ToString = "'-12345'" Then sb.Append("Host_ID,")
        If sInsUpdGuest_ID.ToString = "'-12345'" Then sb.Append("Guest_ID,")
        If sInsUpdArriveDate.ToString = "'Select'" Then sb.Append("ArriveDate,")
        If sInsUpdDepartDate.ToString = "'Select'" Then sb.Append("DepartDate,")
        If sInsUpdStatus.ToString = "'Select'" Then sb.Append("Status,")
        If sInsUpdHostTelephone.ToString = "'Select'" Then sb.Append("HostTelephone,")
        If sInsUpdRate.ToString = "'-12345'" Then sb.Append("Rate,")
        If sInsUpdRequestReceived.ToString = "'Select'" Then sb.Append("RequestReceived,")
        If sInsUpdHostPaidAmtDep.ToString = "'-12345'" Then sb.Append("HostPaidAmtDep,")
        If sInsUpdHostPaidAmtDepFromCredit.ToString = "'-12345'" Then sb.Append("HostPaidAmtDepFromCredit,")
        If sInsUpdDepositDue.ToString = "'Select'" Then sb.Append("DepositDue,")
        If sInsUpdDepositReceived.ToString = "'Select'" Then sb.Append("DepositReceived,")
        If sInsUpdDepositAmount.ToString = "'-12345'" Then sb.Append("DepositAmount,")
        If sInsUpdBalanceDueDate.ToString = "'Select'" Then sb.Append("BalanceDueDate,")
        If sInsUpdBalanceDue.ToString = "'-12345'" Then sb.Append("BalanceDue,")
        If sInsUpdConfirmationSent.ToString = "'Select'" Then sb.Append("ConfirmationSent,")
        If sInsUpdBalanceReceived.ToString = "'Select'" Then sb.Append("BalanceReceived,")
        If sInsUpdBalanceAmount.ToString = "'-12345'" Then sb.Append("BalanceAmount,")
        If sInsUpdHostPdDate.ToString = "'Select'" Then sb.Append("HostPdDate,")
        If sInsUpdHostPaidAmtBal.ToString = "'-12345'" Then sb.Append("HostPaidAmtBal,")
        If sInsUpdHostPaidAmtBalFromCredit.ToString = "'-12345'" Then sb.Append("HostPaidAmtBalFromCredit,")
        If sInsUpdCommission.ToString = "'-12345'" Then sb.Append("Commission,")
        If sInsUpdBookingNotes.ToString = "'Select'" Then sb.Append("BookingNotes,")
        If sInsUpdDepCardType.ToString = "'Select'" Then sb.Append("DepCardType,")
        If sInsUpdDepCardNumber.ToString = "'Select'" Then sb.Append("DepCardNumber,")
        If sInsUpdDepCardNumberEncrypted.ToString = "'Select'" Then sb.Append("DepCardNumberEncrypted,")
        If sInsUpdDepCardConfirm.ToString = "'Select'" Then sb.Append("DepCardConfirm,")
        If sInsUpdBalCardType.ToString = "'Select'" Then sb.Append("BalCardType,")
        If sInsUpdBalCardNumber.ToString = "'Select'" Then sb.Append("BalCardNumber,")
        If sInsUpdBalCardNumberEncrypted.ToString = "'Select'" Then sb.Append("BalCardNumberEncrypted,")
        If sInsUpdBalCardConfirm.ToString = "'Select'" Then sb.Append("BalCardConfirm,")
        If sInsUpdAutoCharge.ToString = "'-12345'" Then sb.Append("AutoCharge,")
        If sInsUpdAdults.ToString = "'-12345'" Then sb.Append("Adults,")
        If sInsUpdChildren.ToString = "'-12345'" Then sb.Append("Children,")
        If sInsUpdTeens.ToString = "'-12345'" Then sb.Append("Teens,")
        If sInsUpdTaxDue.ToString = "'-12345'" Then sb.Append("TaxDue,")
        If sInsUpdRentalBalance.ToString = "'-12345'" Then sb.Append("RentalBalance,")
        If sInsUpdDepCheckName.ToString = "'Select'" Then sb.Append("DepCheckName,")
        If sInsUpdDepCheckNumber.ToString = "'Select'" Then sb.Append("DepCheckNumber,")
        If sInsUpdBalCheckName.ToString = "'Select'" Then sb.Append("BalCheckName,")
        If sInsUpdBalCheckNumber.ToString = "'Select'" Then sb.Append("BalCheckNumber,")
        If sInsUpdAddDate.ToString = "'Select'" Then sb.Append("AddDate,")
        If sInsUpdLocationBooked.ToString = "'Select'" Then sb.Append("LocationBooked,")
        If sInsUpdOrigDepCardNumber.ToString = "'Select'" Then sb.Append("OrigDepCardNumber,")
        If sInsUpdOrigBalCardNumber.ToString = "'Select'" Then sb.Append("OrigBalCardNumber,")
        If sInsUpdCancellationHostPaymentProcessed.ToString = "'-12345'" Then sb.Append("CancellationHostPaymentProcessed,")
        If sInsUpdLateCancellationRefundDate.ToString = "'Select'" Then sb.Append("LateCancellationRefundDate,")
        If sInsUpdLateCancellationRefundStatus.ToString = "'Select'" Then sb.Append("LateCancellationRefundStatus,")
        If sInsUpdLateCancellationRefundAmount.ToString = "'-12345'" Then sb.Append("LateCancellationRefundAmount,")
        If sInsUpdLateCancellationFeeAmount.ToString = "'-12345'" Then sb.Append("LateCancellationFeeAmount,")
        If sInsUpdLateCancellationRefundReBookingID.ToString = "'-12345'" Then sb.Append("LateCancellationRefundReBookingID,")
        If sInsUpdCancellationDate.ToString = "'Select'" Then sb.Append("CancellationDate,")
        If sInsUpdFinalBalancePaymentMadeOnline.ToString = "'-12345'" Then sb.Append("FinalBalancePaymentMadeOnline,")
        If sInsUpdFinalDepositPaymentMadeOnline.ToString = "'-12345'" Then sb.Append("FinalDepositPaymentMadeOnline,")
        If sInsUpdDepositAmountPaid.ToString = "'-12345'" Then sb.Append("DepositAmountPaid,")
        If sInsUpdConfirmationEmail.ToString = "'Select'" Then sb.Append("ConfirmationEmail,")
        If sInsUpdConfirmationEmailSentDate.ToString = "'Select'" Then sb.Append("ConfirmationEmailSentDate,")
        If sInsUpdQBActivityIDFirstPayment.ToString = "'-12345'" Then sb.Append("QBActivityIDFirstPayment,")
        If sInsUpdQBActivityIDBalancePayment.ToString = "'-12345'" Then sb.Append("QBActivityIDBalancePayment,")
        If sInsUpdQBActivityIDFullPayment.ToString = "'-12345'" Then sb.Append("QBActivityIDFullPayment,")
        If sInsUpdPrivateNotes.ToString = "'Select'" Then sb.Append("PrivateNotes,")
      Else
        sb.Append("Booking_ID,")
        sb.Append("Property_ID,")
        sb.Append("Host_ID,")
        sb.Append("Guest_ID,")
        sb.Append("ArriveDate,")
        sb.Append("DepartDate,")
        sb.Append("Status,")
        sb.Append("HostTelephone,")
        sb.Append("Rate,")
        sb.Append("RequestReceived,")
        sb.Append("HostPaidAmtDep,")
        sb.Append("HostPaidAmtDepFromCredit,")
        sb.Append("DepositDue,")
        sb.Append("DepositReceived,")
        sb.Append("DepositAmount,")
        sb.Append("BalanceDueDate,")
        sb.Append("BalanceDue,")
        sb.Append("ConfirmationSent,")
        sb.Append("BalanceReceived,")
        sb.Append("BalanceAmount,")
        sb.Append("HostPdDate,")
        sb.Append("HostPaidAmtBal,")
        sb.Append("HostPaidAmtBalFromCredit,")
        sb.Append("Commission,")
        sb.Append("BookingNotes,")
        sb.Append("DepCardType,")
        sb.Append("DepCardNumber,")
        sb.Append("DepCardNumberEncrypted,")
        sb.Append("DepCardConfirm,")
        sb.Append("BalCardType,")
        sb.Append("BalCardNumber,")
        sb.Append("BalCardNumberEncrypted,")
        sb.Append("BalCardConfirm,")
        sb.Append("AutoCharge,")
        sb.Append("Adults,")
        sb.Append("Children,")
        sb.Append("Teens,")
        sb.Append("TaxDue,")
        sb.Append("RentalBalance,")
        sb.Append("DepCheckName,")
        sb.Append("DepCheckNumber,")
        sb.Append("BalCheckName,")
        sb.Append("BalCheckNumber,")
        sb.Append("AddDate,")
        sb.Append("LocationBooked,")
        sb.Append("OrigDepCardNumber,")
        sb.Append("OrigBalCardNumber,")
        sb.Append("CancellationHostPaymentProcessed,")
        sb.Append("LateCancellationRefundDate,")
        sb.Append("LateCancellationRefundStatus,")
        sb.Append("LateCancellationRefundAmount,")
        sb.Append("LateCancellationFeeAmount,")
        sb.Append("LateCancellationRefundReBookingID,")
        sb.Append("CancellationDate,")
        sb.Append("FinalBalancePaymentMadeOnline,")
        sb.Append("FinalDepositPaymentMadeOnline,")
        sb.Append("DepositAmountPaid,")
        sb.Append("ConfirmationEmail,")
        sb.Append("ConfirmationEmailSentDate,")
        sb.Append("QBActivityIDFirstPayment,")
        sb.Append("QBActivityIDBalancePayment,")
        sb.Append("QBActivityIDFullPayment,")
        sb.Append("PrivateNotes,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [Bookings]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdBooking_ID.ToString <> "") And (sInsUpdBooking_ID <> "'-12345'") Then sbw.Append("Booking_ID=" & sInsUpdBooking_ID & " and ")
      If (sInsUpdProperty_ID.ToString <> "") And (sInsUpdProperty_ID <> "'-12345'") Then sbw.Append("Property_ID=" & sInsUpdProperty_ID & " and ")
      If (sInsUpdHost_ID.ToString <> "") And (sInsUpdHost_ID <> "'-12345'") Then sbw.Append("Host_ID=" & sInsUpdHost_ID & " and ")
      If (sInsUpdGuest_ID.ToString <> "") And (sInsUpdGuest_ID <> "'-12345'") Then sbw.Append("Guest_ID=" & sInsUpdGuest_ID & " and ")
      If (sInsUpdArriveDate.ToString <> "") And (sInsUpdArriveDate <> "'Select'") Then sbw.Append("ArriveDate=" & sInsUpdArriveDate & " and ")
      If (sInsUpdDepartDate.ToString <> "") And (sInsUpdDepartDate <> "'Select'") Then sbw.Append("DepartDate=" & sInsUpdDepartDate & " and ")
      If (sInsUpdStatus.ToString <> "") And (sInsUpdStatus <> "'Select'") Then sbw.Append("Status=" & sInsUpdStatus & " and ")
      If (sInsUpdHostTelephone.ToString <> "") And (sInsUpdHostTelephone <> "'Select'") Then sbw.Append("HostTelephone=" & sInsUpdHostTelephone & " and ")
      If (sInsUpdRate.ToString <> "") And (sInsUpdRate <> "'-12345'") Then sbw.Append("Rate=" & sInsUpdRate & " and ")
      If (sInsUpdRequestReceived.ToString <> "") And (sInsUpdRequestReceived <> "'Select'") Then sbw.Append("RequestReceived=" & sInsUpdRequestReceived & " and ")
      If (sInsUpdHostPaidAmtDep.ToString <> "") And (sInsUpdHostPaidAmtDep <> "'-12345'") Then sbw.Append("HostPaidAmtDep=" & sInsUpdHostPaidAmtDep & " and ")
      If (sInsUpdHostPaidAmtDepFromCredit.ToString <> "") And (sInsUpdHostPaidAmtDepFromCredit <> "'-12345'") Then sbw.Append("HostPaidAmtDepFromCredit=" & sInsUpdHostPaidAmtDepFromCredit & " and ")
      If (sInsUpdDepositDue.ToString <> "") And (sInsUpdDepositDue <> "'Select'") Then sbw.Append("DepositDue=" & sInsUpdDepositDue & " and ")
      If (sInsUpdDepositReceived.ToString <> "") And (sInsUpdDepositReceived <> "'Select'") Then sbw.Append("DepositReceived=" & sInsUpdDepositReceived & " and ")
      If (sInsUpdDepositAmount.ToString <> "") And (sInsUpdDepositAmount <> "'-12345'") Then sbw.Append("DepositAmount=" & sInsUpdDepositAmount & " and ")
      If (sInsUpdBalanceDueDate.ToString <> "") And (sInsUpdBalanceDueDate <> "'Select'") Then sbw.Append("BalanceDueDate=" & sInsUpdBalanceDueDate & " and ")
      If (sInsUpdBalanceDue.ToString <> "") And (sInsUpdBalanceDue <> "'-12345'") Then sbw.Append("BalanceDue=" & sInsUpdBalanceDue & " and ")
      If (sInsUpdConfirmationSent.ToString <> "") And (sInsUpdConfirmationSent <> "'Select'") Then sbw.Append("ConfirmationSent=" & sInsUpdConfirmationSent & " and ")
      If (sInsUpdBalanceReceived.ToString <> "") And (sInsUpdBalanceReceived <> "'Select'") Then sbw.Append("BalanceReceived=" & sInsUpdBalanceReceived & " and ")
      If (sInsUpdBalanceAmount.ToString <> "") And (sInsUpdBalanceAmount <> "'-12345'") Then sbw.Append("BalanceAmount=" & sInsUpdBalanceAmount & " and ")
      If (sInsUpdHostPdDate.ToString <> "") And (sInsUpdHostPdDate <> "'Select'") Then sbw.Append("HostPdDate=" & sInsUpdHostPdDate & " and ")
      If (sInsUpdHostPaidAmtBal.ToString <> "") And (sInsUpdHostPaidAmtBal <> "'-12345'") Then sbw.Append("HostPaidAmtBal=" & sInsUpdHostPaidAmtBal & " and ")
      If (sInsUpdHostPaidAmtBalFromCredit.ToString <> "") And (sInsUpdHostPaidAmtBalFromCredit <> "'-12345'") Then sbw.Append("HostPaidAmtBalFromCredit=" & sInsUpdHostPaidAmtBalFromCredit & " and ")
      If (sInsUpdCommission.ToString <> "") And (sInsUpdCommission <> "'-12345'") Then sbw.Append("Commission=" & sInsUpdCommission & " and ")
      If (sInsUpdBookingNotes.ToString <> "") And (sInsUpdBookingNotes <> "'Select'") Then sbw.Append("BookingNotes=" & sInsUpdBookingNotes & " and ")
      If (sInsUpdDepCardType.ToString <> "") And (sInsUpdDepCardType <> "'Select'") Then sbw.Append("DepCardType=" & sInsUpdDepCardType & " and ")
      If (sInsUpdDepCardNumber.ToString <> "") And (sInsUpdDepCardNumber <> "'Select'") Then sbw.Append("DepCardNumber=" & sInsUpdDepCardNumber & " and ")
      If (sInsUpdDepCardNumberEncrypted.ToString <> "") And (sInsUpdDepCardNumberEncrypted <> "'Select'") Then sbw.Append("DepCardNumberEncrypted=" & sInsUpdDepCardNumberEncrypted & " and ")
      If (sInsUpdDepCardConfirm.ToString <> "") And (sInsUpdDepCardConfirm <> "'Select'") Then sbw.Append("DepCardConfirm=" & sInsUpdDepCardConfirm & " and ")
      If (sInsUpdBalCardType.ToString <> "") And (sInsUpdBalCardType <> "'Select'") Then sbw.Append("BalCardType=" & sInsUpdBalCardType & " and ")
      If (sInsUpdBalCardNumber.ToString <> "") And (sInsUpdBalCardNumber <> "'Select'") Then sbw.Append("BalCardNumber=" & sInsUpdBalCardNumber & " and ")
      If (sInsUpdBalCardNumberEncrypted.ToString <> "") And (sInsUpdBalCardNumberEncrypted <> "'Select'") Then sbw.Append("BalCardNumberEncrypted=" & sInsUpdBalCardNumberEncrypted & " and ")
      If (sInsUpdBalCardConfirm.ToString <> "") And (sInsUpdBalCardConfirm <> "'Select'") Then sbw.Append("BalCardConfirm=" & sInsUpdBalCardConfirm & " and ")
      If (sInsUpdAutoCharge.ToString <> "") And (sInsUpdAutoCharge <> "'-12345'") Then sbw.Append("AutoCharge=" & sInsUpdAutoCharge & " and ")
      If (sInsUpdAdults.ToString <> "") And (sInsUpdAdults <> "'-12345'") Then sbw.Append("Adults=" & sInsUpdAdults & " and ")
      If (sInsUpdChildren.ToString <> "") And (sInsUpdChildren <> "'-12345'") Then sbw.Append("Children=" & sInsUpdChildren & " and ")
      If (sInsUpdTeens.ToString <> "") And (sInsUpdTeens <> "'-12345'") Then sbw.Append("Teens=" & sInsUpdTeens & " and ")
      If (sInsUpdTaxDue.ToString <> "") And (sInsUpdTaxDue <> "'-12345'") Then sbw.Append("TaxDue=" & sInsUpdTaxDue & " and ")
      If (sInsUpdRentalBalance.ToString <> "") And (sInsUpdRentalBalance <> "'-12345'") Then sbw.Append("RentalBalance=" & sInsUpdRentalBalance & " and ")
      If (sInsUpdDepCheckName.ToString <> "") And (sInsUpdDepCheckName <> "'Select'") Then sbw.Append("DepCheckName=" & sInsUpdDepCheckName & " and ")
      If (sInsUpdDepCheckNumber.ToString <> "") And (sInsUpdDepCheckNumber <> "'Select'") Then sbw.Append("DepCheckNumber=" & sInsUpdDepCheckNumber & " and ")
      If (sInsUpdBalCheckName.ToString <> "") And (sInsUpdBalCheckName <> "'Select'") Then sbw.Append("BalCheckName=" & sInsUpdBalCheckName & " and ")
      If (sInsUpdBalCheckNumber.ToString <> "") And (sInsUpdBalCheckNumber <> "'Select'") Then sbw.Append("BalCheckNumber=" & sInsUpdBalCheckNumber & " and ")
      If (sInsUpdAddDate.ToString <> "") And (sInsUpdAddDate <> "'Select'") Then sbw.Append("AddDate=" & sInsUpdAddDate & " and ")
      If (sInsUpdLocationBooked.ToString <> "") And (sInsUpdLocationBooked <> "'Select'") Then sbw.Append("LocationBooked=" & sInsUpdLocationBooked & " and ")
      If (sInsUpdOrigDepCardNumber.ToString <> "") And (sInsUpdOrigDepCardNumber <> "'Select'") Then sbw.Append("OrigDepCardNumber=" & sInsUpdOrigDepCardNumber & " and ")
      If (sInsUpdOrigBalCardNumber.ToString <> "") And (sInsUpdOrigBalCardNumber <> "'Select'") Then sbw.Append("OrigBalCardNumber=" & sInsUpdOrigBalCardNumber & " and ")
      If (sInsUpdCancellationHostPaymentProcessed.ToString <> "") And (sInsUpdCancellationHostPaymentProcessed <> "'-12345'") Then sbw.Append("CancellationHostPaymentProcessed=" & sInsUpdCancellationHostPaymentProcessed & " and ")
      If (sInsUpdLateCancellationRefundDate.ToString <> "") And (sInsUpdLateCancellationRefundDate <> "'Select'") Then sbw.Append("LateCancellationRefundDate=" & sInsUpdLateCancellationRefundDate & " and ")
      If (sInsUpdLateCancellationRefundStatus.ToString <> "") And (sInsUpdLateCancellationRefundStatus <> "'Select'") Then sbw.Append("LateCancellationRefundStatus=" & sInsUpdLateCancellationRefundStatus & " and ")
      If (sInsUpdLateCancellationRefundAmount.ToString <> "") And (sInsUpdLateCancellationRefundAmount <> "'-12345'") Then sbw.Append("LateCancellationRefundAmount=" & sInsUpdLateCancellationRefundAmount & " and ")
      If (sInsUpdLateCancellationFeeAmount.ToString <> "") And (sInsUpdLateCancellationFeeAmount <> "'-12345'") Then sbw.Append("LateCancellationFeeAmount=" & sInsUpdLateCancellationFeeAmount & " and ")
      If (sInsUpdLateCancellationRefundReBookingID.ToString <> "") And (sInsUpdLateCancellationRefundReBookingID <> "'-12345'") Then sbw.Append("LateCancellationRefundReBookingID=" & sInsUpdLateCancellationRefundReBookingID & " and ")
      If (sInsUpdCancellationDate.ToString <> "") And (sInsUpdCancellationDate <> "'Select'") Then sbw.Append("CancellationDate=" & sInsUpdCancellationDate & " and ")
      If (sInsUpdFinalBalancePaymentMadeOnline.ToString <> "") And (sInsUpdFinalBalancePaymentMadeOnline <> "'-12345'") Then sbw.Append("FinalBalancePaymentMadeOnline=" & sInsUpdFinalBalancePaymentMadeOnline & " and ")
      If (sInsUpdFinalDepositPaymentMadeOnline.ToString <> "") And (sInsUpdFinalDepositPaymentMadeOnline <> "'-12345'") Then sbw.Append("FinalDepositPaymentMadeOnline=" & sInsUpdFinalDepositPaymentMadeOnline & " and ")
      If (sInsUpdDepositAmountPaid.ToString <> "") And (sInsUpdDepositAmountPaid <> "'-12345'") Then sbw.Append("DepositAmountPaid=" & sInsUpdDepositAmountPaid & " and ")
      If (sInsUpdConfirmationEmail.ToString <> "") And (sInsUpdConfirmationEmail <> "'Select'") Then sbw.Append("ConfirmationEmail=" & sInsUpdConfirmationEmail & " and ")
      If (sInsUpdConfirmationEmailSentDate.ToString <> "") And (sInsUpdConfirmationEmailSentDate <> "'Select'") Then sbw.Append("ConfirmationEmailSentDate=" & sInsUpdConfirmationEmailSentDate & " and ")
      If (sInsUpdQBActivityIDFirstPayment.ToString <> "") And (sInsUpdQBActivityIDFirstPayment <> "'-12345'") Then sbw.Append("QBActivityIDFirstPayment=" & sInsUpdQBActivityIDFirstPayment & " and ")
      If (sInsUpdQBActivityIDBalancePayment.ToString <> "") And (sInsUpdQBActivityIDBalancePayment <> "'-12345'") Then sbw.Append("QBActivityIDBalancePayment=" & sInsUpdQBActivityIDBalancePayment & " and ")
      If (sInsUpdQBActivityIDFullPayment.ToString <> "") And (sInsUpdQBActivityIDFullPayment <> "'-12345'") Then sbw.Append("QBActivityIDFullPayment=" & sInsUpdQBActivityIDFullPayment & " and ")
      If (sInsUpdPrivateNotes.ToString <> "") And (sInsUpdPrivateNotes <> "'Select'") Then sbw.Append("PrivateNotes=" & sInsUpdPrivateNotes & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If
    oCmd.CommandTimeout = 120

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      Booking_ID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("Booking_ID")), 0, CurrentRow.Item("Booking_ID").ToString)
      Property_ID__Integer = IIf(IsDBNull(CurrentRow.Item("Property_ID")), 0, CurrentRow.Item("Property_ID"))
      Host_ID__Integer = IIf(IsDBNull(CurrentRow.Item("Host_ID")), 0, CurrentRow.Item("Host_ID"))
      Guest_ID__Integer = IIf(IsDBNull(CurrentRow.Item("Guest_ID")), 0, CurrentRow.Item("Guest_ID"))
      ArriveDate__Date = IIf(IsDBNull(CurrentRow.Item("ArriveDate")), "", CurrentRow.Item("ArriveDate"))
      DepartDate__Date = IIf(IsDBNull(CurrentRow.Item("DepartDate")), "", CurrentRow.Item("DepartDate"))
      Status__String = IIf(IsDBNull(CurrentRow.Item("Status")), "", CurrentRow.Item("Status"))
      HostTelephone__String = IIf(IsDBNull(CurrentRow.Item("HostTelephone")), "", CurrentRow.Item("HostTelephone"))
      Rate__Numeric = IIf(IsDBNull(CurrentRow.Item("Rate")), 0.0, CurrentRow.Item("Rate"))
      RequestReceived__Date = IIf(IsDBNull(CurrentRow.Item("RequestReceived")), "", CurrentRow.Item("RequestReceived"))
      HostPaidAmtDep__Numeric = IIf(IsDBNull(CurrentRow.Item("HostPaidAmtDep")), 0.0, CurrentRow.Item("HostPaidAmtDep"))
      HostPaidAmtDepFromCredit__Numeric = IIf(IsDBNull(CurrentRow.Item("HostPaidAmtDepFromCredit")), 0.0, CurrentRow.Item("HostPaidAmtDepFromCredit"))
      DepositDue__Date = IIf(IsDBNull(CurrentRow.Item("DepositDue")), "", CurrentRow.Item("DepositDue"))
      DepositReceived__Date = IIf(IsDBNull(CurrentRow.Item("DepositReceived")), "", CurrentRow.Item("DepositReceived"))
      DepositAmount__Numeric = IIf(IsDBNull(CurrentRow.Item("DepositAmount")), 0.0, CurrentRow.Item("DepositAmount"))
      BalanceDueDate__Date = IIf(IsDBNull(CurrentRow.Item("BalanceDueDate")), "", CurrentRow.Item("BalanceDueDate"))
      BalanceDue__Numeric = IIf(IsDBNull(CurrentRow.Item("BalanceDue")), 0.0, CurrentRow.Item("BalanceDue"))
      ConfirmationSent__Date = IIf(IsDBNull(CurrentRow.Item("ConfirmationSent")), "", CurrentRow.Item("ConfirmationSent"))
      BalanceReceived__Date = IIf(IsDBNull(CurrentRow.Item("BalanceReceived")), "", CurrentRow.Item("BalanceReceived"))
      BalanceAmount__Numeric = IIf(IsDBNull(CurrentRow.Item("BalanceAmount")), 0.0, CurrentRow.Item("BalanceAmount"))
      HostPdDate__Date = IIf(IsDBNull(CurrentRow.Item("HostPdDate")), "", CurrentRow.Item("HostPdDate"))
      HostPaidAmtBal__Numeric = IIf(IsDBNull(CurrentRow.Item("HostPaidAmtBal")), 0.0, CurrentRow.Item("HostPaidAmtBal"))
      HostPaidAmtBalFromCredit__Numeric = IIf(IsDBNull(CurrentRow.Item("HostPaidAmtBalFromCredit")), 0.0, CurrentRow.Item("HostPaidAmtBalFromCredit"))
      Commission__Numeric = IIf(IsDBNull(CurrentRow.Item("Commission")), 0.0, CurrentRow.Item("Commission"))
      BookingNotes__String = IIf(IsDBNull(CurrentRow.Item("BookingNotes")), "", CurrentRow.Item("BookingNotes"))
      DepCardType__String = IIf(IsDBNull(CurrentRow.Item("DepCardType")), "", CurrentRow.Item("DepCardType"))
      DepCardNumber__String = IIf(IsDBNull(CurrentRow.Item("DepCardNumber")), "", CurrentRow.Item("DepCardNumber"))
      DepCardNumberEncrypted__String = IIf(IsDBNull(CurrentRow.Item("DepCardNumberEncrypted")), "", CurrentRow.Item("DepCardNumberEncrypted"))
      DepCardConfirm__String = IIf(IsDBNull(CurrentRow.Item("DepCardConfirm")), "", CurrentRow.Item("DepCardConfirm"))
      BalCardType__String = IIf(IsDBNull(CurrentRow.Item("BalCardType")), "", CurrentRow.Item("BalCardType"))
      BalCardNumber__String = IIf(IsDBNull(CurrentRow.Item("BalCardNumber")), "", CurrentRow.Item("BalCardNumber"))
      BalCardNumberEncrypted__String = IIf(IsDBNull(CurrentRow.Item("BalCardNumberEncrypted")), "", CurrentRow.Item("BalCardNumberEncrypted"))
      BalCardConfirm__String = IIf(IsDBNull(CurrentRow.Item("BalCardConfirm")), "", CurrentRow.Item("BalCardConfirm"))
      AutoCharge__Integer = IIf(IsDBNull(CurrentRow.Item("AutoCharge")), 0.0, CurrentRow.Item("AutoCharge"))
      Adults__Integer = IIf(IsDBNull(CurrentRow.Item("Adults")), 0, CurrentRow.Item("Adults"))
      Children__Integer = IIf(IsDBNull(CurrentRow.Item("Children")), 0, CurrentRow.Item("Children"))
      Teens__Integer = IIf(IsDBNull(CurrentRow.Item("Teens")), 0, CurrentRow.Item("Teens"))
      TaxDue__Numeric = IIf(IsDBNull(CurrentRow.Item("TaxDue")), 0.0, CurrentRow.Item("TaxDue"))
      RentalBalance__Numeric = IIf(IsDBNull(CurrentRow.Item("RentalBalance")), 0.0, CurrentRow.Item("RentalBalance"))
      DepCheckName__String = IIf(IsDBNull(CurrentRow.Item("DepCheckName")), "", CurrentRow.Item("DepCheckName"))
      DepCheckNumber__String = IIf(IsDBNull(CurrentRow.Item("DepCheckNumber")), "", CurrentRow.Item("DepCheckNumber"))
      BalCheckName__String = IIf(IsDBNull(CurrentRow.Item("BalCheckName")), "", CurrentRow.Item("BalCheckName"))
      BalCheckNumber__String = IIf(IsDBNull(CurrentRow.Item("BalCheckNumber")), "", CurrentRow.Item("BalCheckNumber"))
      AddDate__Date = IIf(IsDBNull(CurrentRow.Item("AddDate")), "", CurrentRow.Item("AddDate"))
      LocationBooked__String = IIf(IsDBNull(CurrentRow.Item("LocationBooked")), "", CurrentRow.Item("LocationBooked"))
      OrigDepCardNumber__String = IIf(IsDBNull(CurrentRow.Item("OrigDepCardNumber")), "", CurrentRow.Item("OrigDepCardNumber"))
      OrigBalCardNumber__String = IIf(IsDBNull(CurrentRow.Item("OrigBalCardNumber")), "", CurrentRow.Item("OrigBalCardNumber"))
      CancellationHostPaymentProcessed__Integer = IIf(IsDBNull(CurrentRow.Item("CancellationHostPaymentProcessed")), 0, CurrentRow.Item("CancellationHostPaymentProcessed"))
      LateCancellationRefundDate__Date = IIf(IsDBNull(CurrentRow.Item("LateCancellationRefundDate")), "", CurrentRow.Item("LateCancellationRefundDate"))
      LateCancellationRefundStatus__String = IIf(IsDBNull(CurrentRow.Item("LateCancellationRefundStatus")), "", CurrentRow.Item("LateCancellationRefundStatus"))
      LateCancellationRefundAmount__Numeric = IIf(IsDBNull(CurrentRow.Item("LateCancellationRefundAmount")), 0.0, CurrentRow.Item("LateCancellationRefundAmount"))
      LateCancellationFeeAmount__Numeric = IIf(IsDBNull(CurrentRow.Item("LateCancellationFeeAmount")), 0.0, CurrentRow.Item("LateCancellationFeeAmount"))
      LateCancellationRefundReBookingID__Integer = IIf(IsDBNull(CurrentRow.Item("LateCancellationRefundReBookingID")), 0, CurrentRow.Item("LateCancellationRefundReBookingID"))
      CancellationDate__Date = IIf(IsDBNull(CurrentRow.Item("CancellationDate")), "", CurrentRow.Item("CancellationDate"))
      FinalBalancePaymentMadeOnline__Integer = IIf(IsDBNull(CurrentRow.Item("FinalBalancePaymentMadeOnline")), 0, CurrentRow.Item("FinalBalancePaymentMadeOnline"))
      FinalDepositPaymentMadeOnline__Integer = IIf(IsDBNull(CurrentRow.Item("FinalDepositPaymentMadeOnline")), 0, CurrentRow.Item("FinalDepositPaymentMadeOnline"))
      DepositAmountPaid__Numeric = IIf(IsDBNull(CurrentRow.Item("DepositAmountPaid")), 0.0, CurrentRow.Item("DepositAmountPaid"))
      ConfirmationEmail__String = IIf(IsDBNull(CurrentRow.Item("ConfirmationEmail")), "", CurrentRow.Item("ConfirmationEmail"))
      ConfirmationEmailSentDate__Date = IIf(IsDBNull(CurrentRow.Item("ConfirmationEmailSentDate")), "", CurrentRow.Item("ConfirmationEmailSentDate"))
      QBActivityIDFirstPayment__Integer = IIf(IsDBNull(CurrentRow.Item("QBActivityIDFirstPayment")), 0, CurrentRow.Item("QBActivityIDFirstPayment"))
      QBActivityIDBalancePayment__Integer = IIf(IsDBNull(CurrentRow.Item("QBActivityIDBalancePayment")), 0, CurrentRow.Item("QBActivityIDBalancePayment"))
      QBActivityIDFullPayment__Integer = IIf(IsDBNull(CurrentRow.Item("QBActivityIDFullPayment")), 0, CurrentRow.Item("QBActivityIDFullPayment"))
      PrivateNotes__String = IIf(IsDBNull(CurrentRow.Item("PrivateNotes")), "", CurrentRow.Item("PrivateNotes"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    oCmd.CommandTimeout = 120

    Insert = 0
    sb.Append("Insert into [Bookings](")
    If sInsUpdProperty_ID.ToString <> "" Then
      sb.Append("Property_ID,")
      sbv.Append(sInsUpdProperty_ID & ",")
    End If
    If sInsUpdHost_ID.ToString <> "" Then
      sb.Append("Host_ID,")
      sbv.Append(sInsUpdHost_ID & ",")
    End If
    If sInsUpdGuest_ID.ToString <> "" Then
      sb.Append("Guest_ID,")
      sbv.Append(sInsUpdGuest_ID & ",")
    End If
    If sInsUpdArriveDate.ToString <> "" Then
      sb.Append("ArriveDate,")
      sbv.Append(sInsUpdArriveDate & ",")
    End If
    If sInsUpdDepartDate.ToString <> "" Then
      sb.Append("DepartDate,")
      sbv.Append(sInsUpdDepartDate & ",")
    End If
    If sInsUpdStatus.ToString <> "" Then
      sb.Append("Status,")
      sbv.Append(sInsUpdStatus & ",")
    End If
    If sInsUpdHostTelephone.ToString <> "" Then
      sb.Append("HostTelephone,")
      sbv.Append(sInsUpdHostTelephone & ",")
    End If
    If sInsUpdRate.ToString <> "" Then
      sb.Append("Rate,")
      sbv.Append(sInsUpdRate & ",")
    End If
    If sInsUpdRequestReceived.ToString <> "" Then
      sb.Append("RequestReceived,")
      sbv.Append(sInsUpdRequestReceived & ",")
    End If
    If sInsUpdHostPaidAmtDep.ToString <> "" Then
      sb.Append("HostPaidAmtDep,")
      sbv.Append(sInsUpdHostPaidAmtDep & ",")
    End If
    If sInsUpdHostPaidAmtDepFromCredit.ToString <> "" Then
      sb.Append("HostPaidAmtDepFromCredit,")
      sbv.Append(sInsUpdHostPaidAmtDepFromCredit & ",")
    End If
    If sInsUpdDepositDue.ToString <> "" Then
      sb.Append("DepositDue,")
      sbv.Append(sInsUpdDepositDue & ",")
    End If
    If sInsUpdDepositReceived.ToString <> "" Then
      sb.Append("DepositReceived,")
      sbv.Append(sInsUpdDepositReceived & ",")
    End If
    If sInsUpdDepositAmount.ToString <> "" Then
      sb.Append("DepositAmount,")
      sbv.Append(sInsUpdDepositAmount & ",")
    End If
    If sInsUpdBalanceDueDate.ToString <> "" Then
      sb.Append("BalanceDueDate,")
      sbv.Append(sInsUpdBalanceDueDate & ",")
    End If
    If sInsUpdBalanceDue.ToString <> "" Then
      sb.Append("BalanceDue,")
      sbv.Append(sInsUpdBalanceDue & ",")
    End If
    If sInsUpdConfirmationSent.ToString <> "" Then
      sb.Append("ConfirmationSent,")
      sbv.Append(sInsUpdConfirmationSent & ",")
    End If
    If sInsUpdBalanceReceived.ToString <> "" Then
      sb.Append("BalanceReceived,")
      sbv.Append(sInsUpdBalanceReceived & ",")
    End If
    If sInsUpdBalanceAmount.ToString <> "" Then
      sb.Append("BalanceAmount,")
      sbv.Append(sInsUpdBalanceAmount & ",")
    End If
    If sInsUpdHostPdDate.ToString <> "" Then
      sb.Append("HostPdDate,")
      sbv.Append(sInsUpdHostPdDate & ",")
    End If
    If sInsUpdHostPaidAmtBal.ToString <> "" Then
      sb.Append("HostPaidAmtBal,")
      sbv.Append(sInsUpdHostPaidAmtBal & ",")
    End If
    If sInsUpdHostPaidAmtBalFromCredit.ToString <> "" Then
      sb.Append("HostPaidAmtBalFromCredit,")
      sbv.Append(sInsUpdHostPaidAmtBalFromCredit & ",")
    End If
    If sInsUpdCommission.ToString <> "" Then
      sb.Append("Commission,")
      sbv.Append(sInsUpdCommission & ",")
    End If
    If sInsUpdBookingNotes.ToString <> "" Then
      sb.Append("BookingNotes,")
      sbv.Append(sInsUpdBookingNotes & ",")
    End If
    If sInsUpdDepCardType.ToString <> "" Then
      sb.Append("DepCardType,")
      sbv.Append(sInsUpdDepCardType & ",")
    End If
    If sInsUpdDepCardNumber.ToString <> "" Then
      sb.Append("DepCardNumber,")
      sbv.Append(sInsUpdDepCardNumber & ",")
    End If
    If sInsUpdDepCardNumberEncrypted.ToString <> "" Then
      sb.Append("DepCardNumberEncrypted,")
      sbv.Append(sInsUpdDepCardNumberEncrypted & ",")
    End If
    If sInsUpdDepCardConfirm.ToString <> "" Then
      sb.Append("DepCardConfirm,")
      sbv.Append(sInsUpdDepCardConfirm & ",")
    End If
    If sInsUpdBalCardType.ToString <> "" Then
      sb.Append("BalCardType,")
      sbv.Append(sInsUpdBalCardType & ",")
    End If
    If sInsUpdBalCardNumber.ToString <> "" Then
      sb.Append("BalCardNumber,")
      sbv.Append(sInsUpdBalCardNumber & ",")
    End If
    If sInsUpdBalCardNumberEncrypted.ToString <> "" Then
      sb.Append("BalCardNumberEncrypted,")
      sbv.Append(sInsUpdBalCardNumberEncrypted & ",")
    End If
    If sInsUpdBalCardConfirm.ToString <> "" Then
      sb.Append("BalCardConfirm,")
      sbv.Append(sInsUpdBalCardConfirm & ",")
    End If
    If sInsUpdAutoCharge.ToString <> "" Then
      sb.Append("AutoCharge,")
      sbv.Append(sInsUpdAutoCharge & ",")
    End If
    If sInsUpdAdults.ToString <> "" Then
      sb.Append("Adults,")
      sbv.Append(sInsUpdAdults & ",")
    End If
    If sInsUpdChildren.ToString <> "" Then
      sb.Append("Children,")
      sbv.Append(sInsUpdChildren & ",")
    End If
    If sInsUpdTeens.ToString <> "" Then
      sb.Append("Teens,")
      sbv.Append(sInsUpdTeens & ",")
    End If
    If sInsUpdTaxDue.ToString <> "" Then
      sb.Append("TaxDue,")
      sbv.Append(sInsUpdTaxDue & ",")
    End If
    If sInsUpdRentalBalance.ToString <> "" Then
      sb.Append("RentalBalance,")
      sbv.Append(sInsUpdRentalBalance & ",")
    End If
    If sInsUpdDepCheckName.ToString <> "" Then
      sb.Append("DepCheckName,")
      sbv.Append(sInsUpdDepCheckName & ",")
    End If
    If sInsUpdDepCheckNumber.ToString <> "" Then
      sb.Append("DepCheckNumber,")
      sbv.Append(sInsUpdDepCheckNumber & ",")
    End If
    If sInsUpdBalCheckName.ToString <> "" Then
      sb.Append("BalCheckName,")
      sbv.Append(sInsUpdBalCheckName & ",")
    End If
    If sInsUpdBalCheckNumber.ToString <> "" Then
      sb.Append("BalCheckNumber,")
      sbv.Append(sInsUpdBalCheckNumber & ",")
    End If
    If sInsUpdAddDate.ToString <> "" Then
      sb.Append("AddDate,")
      sbv.Append(sInsUpdAddDate & ",")
    End If
    If sInsUpdLocationBooked.ToString <> "" Then
      sb.Append("LocationBooked,")
      sbv.Append(sInsUpdLocationBooked & ",")
    End If
    If sInsUpdOrigDepCardNumber.ToString <> "" Then
      sb.Append("OrigDepCardNumber,")
      sbv.Append(sInsUpdOrigDepCardNumber & ",")
    End If
    If sInsUpdOrigBalCardNumber.ToString <> "" Then
      sb.Append("OrigBalCardNumber,")
      sbv.Append(sInsUpdOrigBalCardNumber & ",")
    End If
    If sInsUpdCancellationHostPaymentProcessed.ToString <> "" Then
      sb.Append("CancellationHostPaymentProcessed,")
      sbv.Append(sInsUpdCancellationHostPaymentProcessed & ",")
    End If
    If sInsUpdLateCancellationRefundDate.ToString <> "" Then
      sb.Append("LateCancellationRefundDate,")
      sbv.Append(sInsUpdLateCancellationRefundDate & ",")
    End If
    If sInsUpdLateCancellationRefundStatus.ToString <> "" Then
      sb.Append("LateCancellationRefundStatus,")
      sbv.Append(sInsUpdLateCancellationRefundStatus & ",")
    End If
    If sInsUpdLateCancellationRefundAmount.ToString <> "" Then
      sb.Append("LateCancellationRefundAmount,")
      sbv.Append(sInsUpdLateCancellationRefundAmount & ",")
    End If
    If sInsUpdLateCancellationFeeAmount.ToString <> "" Then
      sb.Append("LateCancellationFeeAmount,")
      sbv.Append(sInsUpdLateCancellationFeeAmount & ",")
    End If
    If sInsUpdLateCancellationRefundReBookingID.ToString <> "" Then
      sb.Append("LateCancellationRefundReBookingID,")
      sbv.Append(sInsUpdLateCancellationRefundReBookingID & ",")
    End If
    If sInsUpdCancellationDate.ToString <> "" Then
      sb.Append("CancellationDate,")
      sbv.Append(sInsUpdCancellationDate & ",")
    End If
    If sInsUpdFinalBalancePaymentMadeOnline.ToString <> "" Then
      sb.Append("FinalBalancePaymentMadeOnline,")
      sbv.Append(sInsUpdFinalBalancePaymentMadeOnline & ",")
    End If
    If sInsUpdFinalDepositPaymentMadeOnline.ToString <> "" Then
      sb.Append("FinalDepositPaymentMadeOnline,")
      sbv.Append(sInsUpdFinalDepositPaymentMadeOnline & ",")
    End If
    If sInsUpdDepositAmountPaid.ToString <> "" Then
      sb.Append("DepositAmountPaid,")
      sbv.Append(sInsUpdDepositAmountPaid & ",")
    End If
    If sInsUpdConfirmationEmail.ToString <> "" Then
      sb.Append("ConfirmationEmail,")
      sbv.Append(sInsUpdConfirmationEmail & ",")
    End If
    If sInsUpdConfirmationEmailSentDate.ToString <> "" Then
      sb.Append("ConfirmationEmailSentDate,")
      sbv.Append(sInsUpdConfirmationEmailSentDate & ",")
    End If
    If sInsUpdQBActivityIDFirstPayment.ToString <> "" Then
      sb.Append("QBActivityIDFirstPayment,")
      sbv.Append(sInsUpdQBActivityIDFirstPayment & ",")
    End If
    If sInsUpdQBActivityIDBalancePayment.ToString <> "" Then
      sb.Append("QBActivityIDBalancePayment,")
      sbv.Append(sInsUpdQBActivityIDBalancePayment & ",")
    End If
    If sInsUpdQBActivityIDFullPayment.ToString <> "" Then
      sb.Append("QBActivityIDFullPayment,")
      sbv.Append(sInsUpdQBActivityIDFullPayment & ",")
    End If
    If sInsUpdPrivateNotes.ToString <> "" Then
      sb.Append("PrivateNotes,")
      sbv.Append(sInsUpdPrivateNotes & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(Booking_ID) from [Bookings]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    Booking_ID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    oCmd.CommandTimeout = 120

    Update = 0
    sb.Append("Update [Bookings] Set ")
    If sInsUpdProperty_ID.ToString <> "" Then sb.Append("Property_ID=" & sInsUpdProperty_ID & ",")
    If sInsUpdHost_ID.ToString <> "" Then sb.Append("Host_ID=" & sInsUpdHost_ID & ",")
    If sInsUpdGuest_ID.ToString <> "" Then sb.Append("Guest_ID=" & sInsUpdGuest_ID & ",")
    If sInsUpdArriveDate.ToString <> "" Then sb.Append("ArriveDate=" & sInsUpdArriveDate & ",")
    If sInsUpdDepartDate.ToString <> "" Then sb.Append("DepartDate=" & sInsUpdDepartDate & ",")
    If sInsUpdStatus.ToString <> "" Then sb.Append("Status=" & sInsUpdStatus & ",")
    If sInsUpdHostTelephone.ToString <> "" Then sb.Append("HostTelephone=" & sInsUpdHostTelephone & ",")
    If sInsUpdRate.ToString <> "" Then sb.Append("Rate=" & sInsUpdRate & ",")
    If sInsUpdRequestReceived.ToString <> "" Then sb.Append("RequestReceived=" & sInsUpdRequestReceived & ",")
    If sInsUpdHostPaidAmtDep.ToString <> "" Then sb.Append("HostPaidAmtDep=" & sInsUpdHostPaidAmtDep & ",")
    If sInsUpdHostPaidAmtDepFromCredit.ToString <> "" Then sb.Append("HostPaidAmtDepFromCredit=" & sInsUpdHostPaidAmtDepFromCredit & ",")
    If sInsUpdDepositDue.ToString <> "" Then sb.Append("DepositDue=" & sInsUpdDepositDue & ",")
    If sInsUpdDepositReceived.ToString <> "" Then sb.Append("DepositReceived=" & sInsUpdDepositReceived & ",")
    If sInsUpdDepositAmount.ToString <> "" Then sb.Append("DepositAmount=" & sInsUpdDepositAmount & ",")
    If sInsUpdBalanceDueDate.ToString <> "" Then sb.Append("BalanceDueDate=" & sInsUpdBalanceDueDate & ",")
    If sInsUpdBalanceDue.ToString <> "" Then sb.Append("BalanceDue=" & sInsUpdBalanceDue & ",")
    If sInsUpdConfirmationSent.ToString <> "" Then sb.Append("ConfirmationSent=" & sInsUpdConfirmationSent & ",")
    If sInsUpdBalanceReceived.ToString <> "" Then sb.Append("BalanceReceived=" & sInsUpdBalanceReceived & ",")
    If sInsUpdBalanceAmount.ToString <> "" Then sb.Append("BalanceAmount=" & sInsUpdBalanceAmount & ",")
    If sInsUpdHostPdDate.ToString <> "" Then sb.Append("HostPdDate=" & sInsUpdHostPdDate & ",")
    If sInsUpdHostPaidAmtBal.ToString <> "" Then sb.Append("HostPaidAmtBal=" & sInsUpdHostPaidAmtBal & ",")
    If sInsUpdHostPaidAmtBalFromCredit.ToString <> "" Then sb.Append("HostPaidAmtBalFromCredit=" & sInsUpdHostPaidAmtBalFromCredit & ",")
    If sInsUpdCommission.ToString <> "" Then sb.Append("Commission=" & sInsUpdCommission & ",")
    If sInsUpdBookingNotes.ToString <> "" Then sb.Append("BookingNotes=" & sInsUpdBookingNotes & ",")
    If sInsUpdDepCardType.ToString <> "" Then sb.Append("DepCardType=" & sInsUpdDepCardType & ",")
    If sInsUpdDepCardNumber.ToString <> "" Then sb.Append("DepCardNumber=" & sInsUpdDepCardNumber & ",")
    If sInsUpdDepCardNumberEncrypted.ToString <> "" Then sb.Append("DepCardNumberEncrypted=" & sInsUpdDepCardNumberEncrypted & ",")
    If sInsUpdDepCardConfirm.ToString <> "" Then sb.Append("DepCardConfirm=" & sInsUpdDepCardConfirm & ",")
    If sInsUpdBalCardType.ToString <> "" Then sb.Append("BalCardType=" & sInsUpdBalCardType & ",")
    If sInsUpdBalCardNumber.ToString <> "" Then sb.Append("BalCardNumber=" & sInsUpdBalCardNumber & ",")
    If sInsUpdBalCardNumberEncrypted.ToString <> "" Then sb.Append("BalCardNumberEncrypted=" & sInsUpdBalCardNumberEncrypted & ",")
    If sInsUpdBalCardConfirm.ToString <> "" Then sb.Append("BalCardConfirm=" & sInsUpdBalCardConfirm & ",")
    If sInsUpdAutoCharge.ToString <> "" Then sb.Append("AutoCharge=" & sInsUpdAutoCharge & ",")
    If sInsUpdAdults.ToString <> "" Then sb.Append("Adults=" & sInsUpdAdults & ",")
    If sInsUpdChildren.ToString <> "" Then sb.Append("Children=" & sInsUpdChildren & ",")
    If sInsUpdTeens.ToString <> "" Then sb.Append("Teens=" & sInsUpdTeens & ",")
    If sInsUpdTaxDue.ToString <> "" Then sb.Append("TaxDue=" & sInsUpdTaxDue & ",")
    If sInsUpdRentalBalance.ToString <> "" Then sb.Append("RentalBalance=" & sInsUpdRentalBalance & ",")
    If sInsUpdDepCheckName.ToString <> "" Then sb.Append("DepCheckName=" & sInsUpdDepCheckName & ",")
    If sInsUpdDepCheckNumber.ToString <> "" Then sb.Append("DepCheckNumber=" & sInsUpdDepCheckNumber & ",")
    If sInsUpdBalCheckName.ToString <> "" Then sb.Append("BalCheckName=" & sInsUpdBalCheckName & ",")
    If sInsUpdBalCheckNumber.ToString <> "" Then sb.Append("BalCheckNumber=" & sInsUpdBalCheckNumber & ",")
    If sInsUpdAddDate.ToString <> "" Then sb.Append("AddDate=" & sInsUpdAddDate & ",")
    If sInsUpdLocationBooked.ToString <> "" Then sb.Append("LocationBooked=" & sInsUpdLocationBooked & ",")
    If sInsUpdOrigDepCardNumber.ToString <> "" Then sb.Append("OrigDepCardNumber=" & sInsUpdOrigDepCardNumber & ",")
    If sInsUpdOrigBalCardNumber.ToString <> "" Then sb.Append("OrigBalCardNumber=" & sInsUpdOrigBalCardNumber & ",")
    If sInsUpdCancellationHostPaymentProcessed.ToString <> "" Then sb.Append("CancellationHostPaymentProcessed=" & sInsUpdCancellationHostPaymentProcessed & ",")
    If sInsUpdLateCancellationRefundDate.ToString <> "" Then sb.Append("LateCancellationRefundDate=" & sInsUpdLateCancellationRefundDate & ",")
    If sInsUpdLateCancellationRefundStatus.ToString <> "" Then sb.Append("LateCancellationRefundStatus=" & sInsUpdLateCancellationRefundStatus & ",")
    If sInsUpdLateCancellationRefundAmount.ToString <> "" Then sb.Append("LateCancellationRefundAmount=" & sInsUpdLateCancellationRefundAmount & ",")
    If sInsUpdLateCancellationFeeAmount.ToString <> "" Then sb.Append("LateCancellationFeeAmount=" & sInsUpdLateCancellationFeeAmount & ",")
    If sInsUpdLateCancellationRefundReBookingID.ToString <> "" Then sb.Append("LateCancellationRefundReBookingID=" & sInsUpdLateCancellationRefundReBookingID & ",")
    If sInsUpdCancellationDate.ToString <> "" Then sb.Append("CancellationDate=" & sInsUpdCancellationDate & ",")
    If sInsUpdFinalBalancePaymentMadeOnline.ToString <> "" Then sb.Append("FinalBalancePaymentMadeOnline=" & sInsUpdFinalBalancePaymentMadeOnline & ",")
    If sInsUpdFinalDepositPaymentMadeOnline.ToString <> "" Then sb.Append("FinalDepositPaymentMadeOnline=" & sInsUpdFinalDepositPaymentMadeOnline & ",")
    If sInsUpdDepositAmountPaid.ToString <> "" Then sb.Append("DepositAmountPaid=" & sInsUpdDepositAmountPaid & ",")
    If sInsUpdConfirmationEmail.ToString <> "" Then sb.Append("ConfirmationEmail=" & sInsUpdConfirmationEmail & ",")
    If sInsUpdConfirmationEmailSentDate.ToString <> "" Then sb.Append("ConfirmationEmailSentDate=" & sInsUpdConfirmationEmailSentDate & ",")
    If sInsUpdQBActivityIDFirstPayment.ToString <> "" Then sb.Append("QBActivityIDFirstPayment=" & sInsUpdQBActivityIDFirstPayment & ",")
    If sInsUpdQBActivityIDBalancePayment.ToString <> "" Then sb.Append("QBActivityIDBalancePayment=" & sInsUpdQBActivityIDBalancePayment & ",")
    If sInsUpdQBActivityIDFullPayment.ToString <> "" Then sb.Append("QBActivityIDFullPayment=" & sInsUpdQBActivityIDFullPayment & ",")
    If sInsUpdPrivateNotes.ToString <> "" Then sb.Append("PrivateNotes=" & sInsUpdPrivateNotes & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where Booking_ID=" & sInsUpdBooking_ID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    oCmd.CommandTimeout = 120
    Delete = 0
    sb.Append("Delete [Bookings] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("Booking_ID=" & sInsUpdBooking_ID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TableQBActivity

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iQBActivityID As Int32
  Private sInsUpdQBActivityID As String
  Property ID_PK__Integer() As Int32
    Get
      Return iQBActivityID
    End Get
    Set(ByVal Value As Int32)
      iQBActivityID = Value
      sInsUpdQBActivityID = oUtil.FixParam(iQBActivityID, True)
    End Set
  End Property

  Private sQBActivityDate As String
  Private sInsUpdQBActivityDate As String
  Property Date_RQ__Date() As String
    Get
      Return sQBActivityDate
    End Get
    Set(ByVal Value As String)
      sQBActivityDate = Value
      sInsUpdQBActivityDate = oUtil.FixParam(sQBActivityDate, False)
    End Set
  End Property

  Private sQBActivityType As String
  Private sInsUpdQBActivityType As String
  Property Type_RQ__String() As String
    Get
      Return sQBActivityType
    End Get
    Set(ByVal Value As String)
      sQBActivityType = Value
      sInsUpdQBActivityType = oUtil.FixParam(sQBActivityType, False)
    End Set
  End Property

  Private sQBActivityTxnID As String
  Private sInsUpdQBActivityTxnID As String
  Property TxnID__String() As String
    Get
      Return sQBActivityTxnID
    End Get
    Set(ByVal Value As String)
      sQBActivityTxnID = Value
      sInsUpdQBActivityTxnID = oUtil.FixParam(sQBActivityTxnID, True)
    End Set
  End Property

  Private sQBActivityTxnNumber As String
  Private sInsUpdQBActivityTxnNumber As String
  Property TxnNumber__String() As String
    Get
      Return sQBActivityTxnNumber
    End Get
    Set(ByVal Value As String)
      sQBActivityTxnNumber = Value
      sInsUpdQBActivityTxnNumber = oUtil.FixParam(sQBActivityTxnNumber, True)
    End Set
  End Property

  Private sQBActivityTxnDate As String
  Private sInsUpdQBActivityTxnDate As String
  Property TxnDate__Date() As String
    Get
      Return sQBActivityTxnDate
    End Get
    Set(ByVal Value As String)
      sQBActivityTxnDate = Value
      sInsUpdQBActivityTxnDate = oUtil.FixParam(sQBActivityTxnDate, True)
    End Set
  End Property

  Private sQBActivityTxnTimeCreated As String
  Private sInsUpdQBActivityTxnTimeCreated As String
  Property TxnTimeCreated__Date() As String
    Get
      Return sQBActivityTxnTimeCreated
    End Get
    Set(ByVal Value As String)
      sQBActivityTxnTimeCreated = Value
      sInsUpdQBActivityTxnTimeCreated = oUtil.FixParam(sQBActivityTxnTimeCreated, True)
    End Set
  End Property

  Private dQBActivityAmount As Double
  Private sInsUpdQBActivityAmount As String
  Property Amount__Numeric() As Double
    Get
      Return dQBActivityAmount
    End Get
    Set(ByVal Value As Double)
      dQBActivityAmount = Value
      sInsUpdQBActivityAmount = oUtil.FixParam(dQBActivityAmount, True)
    End Set
  End Property

  Private sQBActivityDepositType As String
  Private sInsUpdQBActivityDepositType As String
  Property DepositType__String() As String
    Get
      Return sQBActivityDepositType
    End Get
    Set(ByVal Value As String)
      sQBActivityDepositType = Value
      sInsUpdQBActivityDepositType = oUtil.FixParam(sQBActivityDepositType, True)
    End Set
  End Property

  Private sQBActivityStatus As String
  Private sInsUpdQBActivityStatus As String
  Property Status__String() As String
    Get
      Return sQBActivityStatus
    End Get
    Set(ByVal Value As String)
      sQBActivityStatus = Value
      sInsUpdQBActivityStatus = oUtil.FixParam(sQBActivityStatus, True)
    End Set
  End Property

  Private iQBActivityBookingID As Int32
  Private sInsUpdQBActivityBookingID As String
  Property BookingID__Integer() As Int32
    Get
      Return iQBActivityBookingID
    End Get
    Set(ByVal Value As Int32)
      iQBActivityBookingID = Value
      sInsUpdQBActivityBookingID = oUtil.FixParam(iQBActivityBookingID, True)
    End Set
  End Property

  Private sQBActivityPaymentIDs As String
  Private sInsUpdQBActivityPaymentIDs As String
  Property PaymentIDs__String() As String
    Get
      Return sQBActivityPaymentIDs
    End Get
    Set(ByVal Value As String)
      sQBActivityPaymentIDs = Value
      sInsUpdQBActivityPaymentIDs = oUtil.FixParam(sQBActivityPaymentIDs, True)
    End Set
  End Property

  Private sQBActivityNotes As String
  Private sInsUpdQBActivityNotes As String
  Property Notes__String() As String
    Get
      Return sQBActivityNotes
    End Get
    Set(ByVal Value As String)
      sQBActivityNotes = Value
      sInsUpdQBActivityNotes = oUtil.FixParam(sQBActivityNotes, True)
    End Set
  End Property

  Private sQBActivityPayee As String
  Private sInsUpdQBActivityPayee As String
  Property Payee__String() As String
    Get
      Return sQBActivityPayee
    End Get
    Set(ByVal Value As String)
      sQBActivityPayee = Value
      sInsUpdQBActivityPayee = oUtil.FixParam(sQBActivityPayee, True)
    End Set
  End Property

  Public Sub Clear()
    iQBActivityID = 0
    sInsUpdQBActivityID = ""
    sQBActivityDate = ""
    sInsUpdQBActivityDate = ""
    sQBActivityType = ""
    sInsUpdQBActivityType = ""
    sQBActivityTxnID = ""
    sInsUpdQBActivityTxnID = ""
    sQBActivityTxnNumber = ""
    sInsUpdQBActivityTxnNumber = ""
    sQBActivityTxnDate = ""
    sInsUpdQBActivityTxnDate = ""
    sQBActivityTxnTimeCreated = ""
    sInsUpdQBActivityTxnTimeCreated = ""
    dQBActivityAmount = 0.0
    sInsUpdQBActivityAmount = ""
    sQBActivityDepositType = ""
    sInsUpdQBActivityDepositType = ""
    sQBActivityStatus = ""
    sInsUpdQBActivityStatus = ""
    iQBActivityBookingID = 0
    sInsUpdQBActivityBookingID = ""
    sQBActivityPaymentIDs = ""
    sInsUpdQBActivityPaymentIDs = ""
    sQBActivityNotes = ""
    sInsUpdQBActivityNotes = ""
    sQBActivityPayee = ""
    sInsUpdQBActivityPayee = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    oCmd.CommandTimeout = 120

    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdQBActivityID.ToString = "'-12345'" Then sb.Append("QBActivityID,")
        If sInsUpdQBActivityDate.ToString = "'Select'" Then sb.Append("QBActivityDate,")
        If sInsUpdQBActivityType.ToString = "'Select'" Then sb.Append("QBActivityType,")
        If sInsUpdQBActivityTxnID.ToString = "'Select'" Then sb.Append("QBActivityTxnID,")
        If sInsUpdQBActivityTxnNumber.ToString = "'Select'" Then sb.Append("QBActivityTxnNumber,")
        If sInsUpdQBActivityTxnDate.ToString = "'Select'" Then sb.Append("QBActivityTxnDate,")
        If sInsUpdQBActivityTxnTimeCreated.ToString = "'Select'" Then sb.Append("QBActivityTxnTimeCreated,")
        If sInsUpdQBActivityAmount.ToString = "'-12345'" Then sb.Append("QBActivityAmount,")
        If sInsUpdQBActivityDepositType.ToString = "'Select'" Then sb.Append("QBActivityDepositType,")
        If sInsUpdQBActivityStatus.ToString = "'Select'" Then sb.Append("QBActivityStatus,")
        If sInsUpdQBActivityBookingID.ToString = "'-12345'" Then sb.Append("QBActivityBookingID,")
        If sInsUpdQBActivityPaymentIDs.ToString = "'Select'" Then sb.Append("QBActivityPaymentIDs,")
        If sInsUpdQBActivityNotes.ToString = "'Select'" Then sb.Append("QBActivityNotes,")
        If sInsUpdQBActivityPayee.ToString = "'Select'" Then sb.Append("QBActivityPayee,")
      Else
        sb.Append("QBActivityID,")
        sb.Append("QBActivityDate,")
        sb.Append("QBActivityType,")
        sb.Append("QBActivityTxnID,")
        sb.Append("QBActivityTxnNumber,")
        sb.Append("QBActivityTxnDate,")
        sb.Append("QBActivityTxnTimeCreated,")
        sb.Append("QBActivityAmount,")
        sb.Append("QBActivityDepositType,")
        sb.Append("QBActivityStatus,")
        sb.Append("QBActivityBookingID,")
        sb.Append("QBActivityPaymentIDs,")
        sb.Append("QBActivityNotes,")
        sb.Append("QBActivityPayee,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [QBActivity]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdQBActivityID.ToString <> "") And (sInsUpdQBActivityID <> "'-12345'") Then sbw.Append("QBActivityID=" & sInsUpdQBActivityID & " and ")
      If (sInsUpdQBActivityDate.ToString <> "") And (sInsUpdQBActivityDate <> "'Select'") Then sbw.Append("QBActivityDate=" & sInsUpdQBActivityDate & " and ")
      If (sInsUpdQBActivityType.ToString <> "") And (sInsUpdQBActivityType <> "'Select'") Then sbw.Append("QBActivityType=" & sInsUpdQBActivityType & " and ")
      If (sInsUpdQBActivityTxnID.ToString <> "") And (sInsUpdQBActivityTxnID <> "'Select'") Then sbw.Append("QBActivityTxnID=" & sInsUpdQBActivityTxnID & " and ")
      If (sInsUpdQBActivityTxnNumber.ToString <> "") And (sInsUpdQBActivityTxnNumber <> "'Select'") Then sbw.Append("QBActivityTxnNumber=" & sInsUpdQBActivityTxnNumber & " and ")
      If (sInsUpdQBActivityTxnDate.ToString <> "") And (sInsUpdQBActivityTxnDate <> "'Select'") Then sbw.Append("QBActivityTxnDate=" & sInsUpdQBActivityTxnDate & " and ")
      If (sInsUpdQBActivityTxnTimeCreated.ToString <> "") And (sInsUpdQBActivityTxnTimeCreated <> "'Select'") Then sbw.Append("QBActivityTxnTimeCreated=" & sInsUpdQBActivityTxnTimeCreated & " and ")
      If (sInsUpdQBActivityAmount.ToString <> "") And (sInsUpdQBActivityAmount <> "'-12345'") Then sbw.Append("QBActivityAmount=" & sInsUpdQBActivityAmount & " and ")
      If (sInsUpdQBActivityDepositType.ToString <> "") And (sInsUpdQBActivityDepositType <> "'Select'") Then sbw.Append("QBActivityDepositType=" & sInsUpdQBActivityDepositType & " and ")
      If (sInsUpdQBActivityStatus.ToString <> "") And (sInsUpdQBActivityStatus <> "'Select'") Then sbw.Append("QBActivityStatus=" & sInsUpdQBActivityStatus & " and ")
      If (sInsUpdQBActivityBookingID.ToString <> "") And (sInsUpdQBActivityBookingID <> "'-12345'") Then sbw.Append("QBActivityBookingID=" & sInsUpdQBActivityBookingID & " and ")
      If (sInsUpdQBActivityPaymentIDs.ToString <> "") And (sInsUpdQBActivityPaymentIDs <> "'Select'") Then sbw.Append("QBActivityPaymentIDs=" & sInsUpdQBActivityPaymentIDs & " and ")
      If (sInsUpdQBActivityNotes.ToString <> "") And (sInsUpdQBActivityNotes <> "'Select'") Then sbw.Append("QBActivityNotes=" & sInsUpdQBActivityNotes & " and ")
      If (sInsUpdQBActivityPayee.ToString <> "") And (sInsUpdQBActivityPayee <> "'Select'") Then sbw.Append("QBActivityPayee=" & sInsUpdQBActivityPayee & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    oCmd.CommandTimeout = 120

    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      ID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("QBActivityID")), 0, CurrentRow.Item("QBActivityID").ToString)
      Date_RQ__Date = IIf(IsDBNull(CurrentRow.Item("QBActivityDate")), "", CurrentRow.Item("QBActivityDate"))
      Type_RQ__String = IIf(IsDBNull(CurrentRow.Item("QBActivityType")), "", CurrentRow.Item("QBActivityType"))
      TxnID__String = IIf(IsDBNull(CurrentRow.Item("QBActivityTxnID")), "", CurrentRow.Item("QBActivityTxnID"))
      TxnNumber__String = IIf(IsDBNull(CurrentRow.Item("QBActivityTxnNumber")), "", CurrentRow.Item("QBActivityTxnNumber"))
      TxnDate__Date = IIf(IsDBNull(CurrentRow.Item("QBActivityTxnDate")), "", CurrentRow.Item("QBActivityTxnDate"))
      TxnTimeCreated__Date = IIf(IsDBNull(CurrentRow.Item("QBActivityTxnTimeCreated")), "", CurrentRow.Item("QBActivityTxnTimeCreated"))
      Amount__Numeric = IIf(IsDBNull(CurrentRow.Item("QBActivityAmount")), 0.0, CurrentRow.Item("QBActivityAmount"))
      DepositType__String = IIf(IsDBNull(CurrentRow.Item("QBActivityDepositType")), "", CurrentRow.Item("QBActivityDepositType"))
      Status__String = IIf(IsDBNull(CurrentRow.Item("QBActivityStatus")), "", CurrentRow.Item("QBActivityStatus"))
      BookingID__Integer = IIf(IsDBNull(CurrentRow.Item("QBActivityBookingID")), 0, CurrentRow.Item("QBActivityBookingID"))
      PaymentIDs__String = IIf(IsDBNull(CurrentRow.Item("QBActivityPaymentIDs")), "", CurrentRow.Item("QBActivityPaymentIDs"))
      Notes__String = IIf(IsDBNull(CurrentRow.Item("QBActivityNotes")), "", CurrentRow.Item("QBActivityNotes"))
      Payee__String = IIf(IsDBNull(CurrentRow.Item("QBActivityPayee")), "", CurrentRow.Item("QBActivityPayee"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    oCmd.CommandTimeout = 120

    Insert = 0
    sb.Append("Insert into [QBActivity](")
    If sInsUpdQBActivityDate.ToString <> "" Then
      sb.Append("QBActivityDate,")
      sbv.Append(sInsUpdQBActivityDate & ",")
    End If
    If sInsUpdQBActivityType.ToString <> "" Then
      sb.Append("QBActivityType,")
      sbv.Append(sInsUpdQBActivityType & ",")
    End If
    If sInsUpdQBActivityTxnID.ToString <> "" Then
      sb.Append("QBActivityTxnID,")
      sbv.Append(sInsUpdQBActivityTxnID & ",")
    End If
    If sInsUpdQBActivityTxnNumber.ToString <> "" Then
      sb.Append("QBActivityTxnNumber,")
      sbv.Append(sInsUpdQBActivityTxnNumber & ",")
    End If
    If sInsUpdQBActivityTxnDate.ToString <> "" Then
      sb.Append("QBActivityTxnDate,")
      sbv.Append(sInsUpdQBActivityTxnDate & ",")
    End If
    If sInsUpdQBActivityTxnTimeCreated.ToString <> "" Then
      sb.Append("QBActivityTxnTimeCreated,")
      sbv.Append(sInsUpdQBActivityTxnTimeCreated & ",")
    End If
    If sInsUpdQBActivityAmount.ToString <> "" Then
      sb.Append("QBActivityAmount,")
      sbv.Append(sInsUpdQBActivityAmount & ",")
    End If
    If sInsUpdQBActivityDepositType.ToString <> "" Then
      sb.Append("QBActivityDepositType,")
      sbv.Append(sInsUpdQBActivityDepositType & ",")
    End If
    If sInsUpdQBActivityStatus.ToString <> "" Then
      sb.Append("QBActivityStatus,")
      sbv.Append(sInsUpdQBActivityStatus & ",")
    End If
    If sInsUpdQBActivityBookingID.ToString <> "" Then
      sb.Append("QBActivityBookingID,")
      sbv.Append(sInsUpdQBActivityBookingID & ",")
    End If
    If sInsUpdQBActivityPaymentIDs.ToString <> "" Then
      sb.Append("QBActivityPaymentIDs,")
      sbv.Append(sInsUpdQBActivityPaymentIDs & ",")
    End If
    If sInsUpdQBActivityNotes.ToString <> "" Then
      sb.Append("QBActivityNotes,")
      sbv.Append(sInsUpdQBActivityNotes & ",")
    End If
    If sInsUpdQBActivityPayee.ToString <> "" Then
      sb.Append("QBActivityPayee,")
      sbv.Append(sInsUpdQBActivityPayee & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(QBActivityID) from [QBActivity]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    ID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [QBActivity] Set ")
    If sInsUpdQBActivityDate.ToString <> "" Then sb.Append("QBActivityDate=" & sInsUpdQBActivityDate & ",")
    If sInsUpdQBActivityType.ToString <> "" Then sb.Append("QBActivityType=" & sInsUpdQBActivityType & ",")
    If sInsUpdQBActivityTxnID.ToString <> "" Then sb.Append("QBActivityTxnID=" & sInsUpdQBActivityTxnID & ",")
    If sInsUpdQBActivityTxnNumber.ToString <> "" Then sb.Append("QBActivityTxnNumber=" & sInsUpdQBActivityTxnNumber & ",")
    If sInsUpdQBActivityTxnDate.ToString <> "" Then sb.Append("QBActivityTxnDate=" & sInsUpdQBActivityTxnDate & ",")
    If sInsUpdQBActivityTxnTimeCreated.ToString <> "" Then sb.Append("QBActivityTxnTimeCreated=" & sInsUpdQBActivityTxnTimeCreated & ",")
    If sInsUpdQBActivityAmount.ToString <> "" Then sb.Append("QBActivityAmount=" & sInsUpdQBActivityAmount & ",")
    If sInsUpdQBActivityDepositType.ToString <> "" Then sb.Append("QBActivityDepositType=" & sInsUpdQBActivityDepositType & ",")
    If sInsUpdQBActivityStatus.ToString <> "" Then sb.Append("QBActivityStatus=" & sInsUpdQBActivityStatus & ",")
    If sInsUpdQBActivityBookingID.ToString <> "" Then sb.Append("QBActivityBookingID=" & sInsUpdQBActivityBookingID & ",")
    If sInsUpdQBActivityPaymentIDs.ToString <> "" Then sb.Append("QBActivityPaymentIDs=" & sInsUpdQBActivityPaymentIDs & ",")
    If sInsUpdQBActivityNotes.ToString <> "" Then sb.Append("QBActivityNotes=" & sInsUpdQBActivityNotes & ",")
    If sInsUpdQBActivityPayee.ToString <> "" Then sb.Append("QBActivityPayee=" & sInsUpdQBActivityPayee & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where QBActivityID=" & sInsUpdQBActivityID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [QBActivity] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("QBActivityID=" & sInsUpdQBActivityID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TableGuests

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iGuest_ID As Int32
  Private sInsUpdGuest_ID As String
  Property Guest_ID_PK__Integer() As Int32
    Get
      Return iGuest_ID
    End Get
    Set(ByVal Value As Int32)
      iGuest_ID = Value
      sInsUpdGuest_ID = oUtil.FixParam(iGuest_ID, True)
    End Set
  End Property

  Private sGuestName As String
  Private sInsUpdGuestName As String
  Property GuestName__String() As String
    Get
      Return sGuestName
    End Get
    Set(ByVal Value As String)
      sGuestName = Value
      sInsUpdGuestName = oUtil.FixParam(sGuestName, True)
    End Set
  End Property

  Private sGuestFirstName As String
  Private sInsUpdGuestFirstName As String
  Property GuestFirstName__String() As String
    Get
      Return sGuestFirstName
    End Get
    Set(ByVal Value As String)
      sGuestFirstName = Value
      sInsUpdGuestFirstName = oUtil.FixParam(sGuestFirstName, True)
    End Set
  End Property

  Private sGuestLastName As String
  Private sInsUpdGuestLastName As String
  Property GuestLastName__String() As String
    Get
      Return sGuestLastName
    End Get
    Set(ByVal Value As String)
      sGuestLastName = Value
      sInsUpdGuestLastName = oUtil.FixParam(sGuestLastName, True)
    End Set
  End Property

  Private sAddress As String
  Private sInsUpdAddress As String
  Property Address__String() As String
    Get
      Return sAddress
    End Get
    Set(ByVal Value As String)
      sAddress = Value
      sInsUpdAddress = oUtil.FixParam(sAddress, True)
    End Set
  End Property

  Private sAddress2 As String
  Private sInsUpdAddress2 As String
  Property Address2__String() As String
    Get
      Return sAddress2
    End Get
    Set(ByVal Value As String)
      sAddress2 = Value
      sInsUpdAddress2 = oUtil.FixParam(sAddress2, True)
    End Set
  End Property

  Private sCity As String
  Private sInsUpdCity As String
  Property City__String() As String
    Get
      Return sCity
    End Get
    Set(ByVal Value As String)
      sCity = Value
      sInsUpdCity = oUtil.FixParam(sCity, True)
    End Set
  End Property

  Private sState As String
  Private sInsUpdState As String
  Property State__String() As String
    Get
      Return sState
    End Get
    Set(ByVal Value As String)
      sState = Value
      sInsUpdState = oUtil.FixParam(sState, True)
    End Set
  End Property

  Private sZip As String
  Private sInsUpdZip As String
  Property Zip__String() As String
    Get
      Return sZip
    End Get
    Set(ByVal Value As String)
      sZip = Value
      sInsUpdZip = oUtil.FixParam(sZip, True)
    End Set
  End Property

  Private sHomePhone As String
  Private sInsUpdHomePhone As String
  Property HomePhone__String() As String
    Get
      Return sHomePhone
    End Get
    Set(ByVal Value As String)
      sHomePhone = Value
      sInsUpdHomePhone = oUtil.FixParam(sHomePhone, True)
    End Set
  End Property

  Private sWorkPhone As String
  Private sInsUpdWorkPhone As String
  Property WorkPhone__String() As String
    Get
      Return sWorkPhone
    End Get
    Set(ByVal Value As String)
      sWorkPhone = Value
      sInsUpdWorkPhone = oUtil.FixParam(sWorkPhone, True)
    End Set
  End Property

  Private sCellPhone As String
  Private sInsUpdCellPhone As String
  Property CellPhone__String() As String
    Get
      Return sCellPhone
    End Get
    Set(ByVal Value As String)
      sCellPhone = Value
      sInsUpdCellPhone = oUtil.FixParam(sCellPhone, True)
    End Set
  End Property

  Private sEmail As String
  Private sInsUpdEmail As String
  Property Email__String() As String
    Get
      Return sEmail
    End Get
    Set(ByVal Value As String)
      sEmail = Value
      sInsUpdEmail = oUtil.FixParam(sEmail, True)
    End Set
  End Property

  Private sGuestNotes As String
  Private sInsUpdGuestNotes As String
  Property GuestNotes__String() As String
    Get
      Return sGuestNotes
    End Get
    Set(ByVal Value As String)
      sGuestNotes = Value
      sInsUpdGuestNotes = oUtil.FixParam(sGuestNotes, True)
    End Set
  End Property

  Private sCellPhone2 As String
  Private sInsUpdCellPhone2 As String
  Property CellPhone2__String() As String
    Get
      Return sCellPhone2
    End Get
    Set(ByVal Value As String)
      sCellPhone2 = Value
      sInsUpdCellPhone2 = oUtil.FixParam(sCellPhone2, True)
    End Set
  End Property

  Private sEmail2 As String
  Private sInsUpdEmail2 As String
  Property Email2__String() As String
    Get
      Return sEmail2
    End Get
    Set(ByVal Value As String)
      sEmail2 = Value
      sInsUpdEmail2 = oUtil.FixParam(sEmail2, True)
    End Set
  End Property

  Private sCountry As String
  Private sInsUpdCountry As String
  Property Country__String() As String
    Get
      Return sCountry
    End Get
    Set(ByVal Value As String)
      sCountry = Value
      sInsUpdCountry = oUtil.FixParam(sCountry, True)
    End Set
  End Property

  Private sQBListID As String
  Private sInsUpdQBListID As String
  Property QBListID__String() As String
    Get
      Return sQBListID
    End Get
    Set(ByVal Value As String)
      sQBListID = Value
      sInsUpdQBListID = oUtil.FixParam(sQBListID, True)
    End Set
  End Property

  Private sAuthNETCustomerProfileID As String
  Private sInsUpdAuthNETCustomerProfileID As String
  Property AuthNETCustomerProfileID__String() As String
    Get
      Return sAuthNETCustomerProfileID
    End Get
    Set(ByVal Value As String)
      sAuthNETCustomerProfileID = Value
      sInsUpdAuthNETCustomerProfileID = oUtil.FixParam(sAuthNETCustomerProfileID, True)
    End Set
  End Property

  Public Sub Clear()
    iGuest_ID = 0
    sInsUpdGuest_ID = ""
    sGuestName = ""
    sInsUpdGuestName = ""
    sGuestFirstName = ""
    sInsUpdGuestFirstName = ""
    sGuestLastName = ""
    sInsUpdGuestLastName = ""
    sAddress = ""
    sInsUpdAddress = ""
    sAddress2 = ""
    sInsUpdAddress2 = ""
    sCity = ""
    sInsUpdCity = ""
    sState = ""
    sInsUpdState = ""
    sZip = ""
    sInsUpdZip = ""
    sHomePhone = ""
    sInsUpdHomePhone = ""
    sWorkPhone = ""
    sInsUpdWorkPhone = ""
    sCellPhone = ""
    sInsUpdCellPhone = ""
    sEmail = ""
    sInsUpdEmail = ""
    sGuestNotes = ""
    sInsUpdGuestNotes = ""
    sCellPhone2 = ""
    sInsUpdCellPhone2 = ""
    sEmail2 = ""
    sInsUpdEmail2 = ""
    sCountry = ""
    sInsUpdCountry = ""
    sQBListID = ""
    sInsUpdQBListID = ""
    sAuthNETCustomerProfileID = ""
    sInsUpdAuthNETCustomerProfileID = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdGuest_ID.ToString = "'-12345'" Then sb.Append("Guest_ID,")
        If sInsUpdGuestName.ToString = "'Select'" Then sb.Append("GuestName,")
        If sInsUpdGuestFirstName.ToString = "'Select'" Then sb.Append("GuestFirstName,")
        If sInsUpdGuestLastName.ToString = "'Select'" Then sb.Append("GuestLastName,")
        If sInsUpdAddress.ToString = "'Select'" Then sb.Append("Address,")
        If sInsUpdAddress2.ToString = "'Select'" Then sb.Append("Address2,")
        If sInsUpdCity.ToString = "'Select'" Then sb.Append("City,")
        If sInsUpdState.ToString = "'Select'" Then sb.Append("State,")
        If sInsUpdZip.ToString = "'Select'" Then sb.Append("Zip,")
        If sInsUpdHomePhone.ToString = "'Select'" Then sb.Append("HomePhone,")
        If sInsUpdWorkPhone.ToString = "'Select'" Then sb.Append("WorkPhone,")
        If sInsUpdCellPhone.ToString = "'Select'" Then sb.Append("CellPhone,")
        If sInsUpdEmail.ToString = "'Select'" Then sb.Append("Email,")
        If sInsUpdGuestNotes.ToString = "'Select'" Then sb.Append("GuestNotes,")
        If sInsUpdCellPhone2.ToString = "'Select'" Then sb.Append("CellPhone2,")
        If sInsUpdEmail2.ToString = "'Select'" Then sb.Append("Email2,")
        If sInsUpdCountry.ToString = "'Select'" Then sb.Append("Country,")
        If sInsUpdQBListID.ToString = "'Select'" Then sb.Append("QBListID,")
        If sInsUpdAuthNETCustomerProfileID.ToString = "'Select'" Then sb.Append("AuthNETCustomerProfileID,")
      Else
        sb.Append("Guest_ID,")
        sb.Append("GuestName,")
        sb.Append("GuestFirstName,")
        sb.Append("GuestLastName,")
        sb.Append("Address,")
        sb.Append("Address2,")
        sb.Append("City,")
        sb.Append("State,")
        sb.Append("Zip,")
        sb.Append("HomePhone,")
        sb.Append("WorkPhone,")
        sb.Append("CellPhone,")
        sb.Append("Email,")
        sb.Append("GuestNotes,")
        sb.Append("CellPhone2,")
        sb.Append("Email2,")
        sb.Append("Country,")
        sb.Append("QBListID,")
        sb.Append("AuthNETCustomerProfileID,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [Guests]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdGuest_ID.ToString <> "") And (sInsUpdGuest_ID <> "'-12345'") Then sbw.Append("Guest_ID=" & sInsUpdGuest_ID & " and ")
      If (sInsUpdGuestName.ToString <> "") And (sInsUpdGuestName <> "'Select'") Then sbw.Append("GuestName=" & sInsUpdGuestName & " and ")
      If (sInsUpdGuestFirstName.ToString <> "") And (sInsUpdGuestFirstName <> "'Select'") Then sbw.Append("GuestFirstName=" & sInsUpdGuestFirstName & " and ")
      If (sInsUpdGuestLastName.ToString <> "") And (sInsUpdGuestLastName <> "'Select'") Then sbw.Append("GuestLastName=" & sInsUpdGuestLastName & " and ")
      If (sInsUpdAddress.ToString <> "") And (sInsUpdAddress <> "'Select'") Then sbw.Append("Address=" & sInsUpdAddress & " and ")
      If (sInsUpdAddress2.ToString <> "") And (sInsUpdAddress2 <> "'Select'") Then sbw.Append("Address2=" & sInsUpdAddress2 & " and ")
      If (sInsUpdCity.ToString <> "") And (sInsUpdCity <> "'Select'") Then sbw.Append("City=" & sInsUpdCity & " and ")
      If (sInsUpdState.ToString <> "") And (sInsUpdState <> "'Select'") Then sbw.Append("State=" & sInsUpdState & " and ")
      If (sInsUpdZip.ToString <> "") And (sInsUpdZip <> "'Select'") Then sbw.Append("Zip=" & sInsUpdZip & " and ")
      If (sInsUpdHomePhone.ToString <> "") And (sInsUpdHomePhone <> "'Select'") Then sbw.Append("HomePhone=" & sInsUpdHomePhone & " and ")
      If (sInsUpdWorkPhone.ToString <> "") And (sInsUpdWorkPhone <> "'Select'") Then sbw.Append("WorkPhone=" & sInsUpdWorkPhone & " and ")
      If (sInsUpdCellPhone.ToString <> "") And (sInsUpdCellPhone <> "'Select'") Then sbw.Append("CellPhone=" & sInsUpdCellPhone & " and ")
      If (sInsUpdEmail.ToString <> "") And (sInsUpdEmail <> "'Select'") Then sbw.Append("Email=" & sInsUpdEmail & " and ")
      If (sInsUpdGuestNotes.ToString <> "") And (sInsUpdGuestNotes <> "'Select'") Then sbw.Append("GuestNotes=" & sInsUpdGuestNotes & " and ")
      If (sInsUpdCellPhone2.ToString <> "") And (sInsUpdCellPhone2 <> "'Select'") Then sbw.Append("CellPhone2=" & sInsUpdCellPhone2 & " and ")
      If (sInsUpdEmail2.ToString <> "") And (sInsUpdEmail2 <> "'Select'") Then sbw.Append("Email2=" & sInsUpdEmail2 & " and ")
      If (sInsUpdCountry.ToString <> "") And (sInsUpdCountry <> "'Select'") Then sbw.Append("Country=" & sInsUpdCountry & " and ")
      If (sInsUpdQBListID.ToString <> "") And (sInsUpdQBListID <> "'Select'") Then sbw.Append("QBListID=" & sInsUpdQBListID & " and ")
      If (sInsUpdAuthNETCustomerProfileID.ToString <> "") And (sInsUpdAuthNETCustomerProfileID <> "'Select'") Then sbw.Append("AuthNETCustomerProfileID=" & sInsUpdAuthNETCustomerProfileID & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      Guest_ID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("Guest_ID")), 0, CurrentRow.Item("Guest_ID").ToString)
      GuestName__String = IIf(IsDBNull(CurrentRow.Item("GuestName")), "", CurrentRow.Item("GuestName"))
      GuestFirstName__String = IIf(IsDBNull(CurrentRow.Item("GuestFirstName")), "", CurrentRow.Item("GuestFirstName"))
      GuestLastName__String = IIf(IsDBNull(CurrentRow.Item("GuestLastName")), "", CurrentRow.Item("GuestLastName"))
      Address__String = IIf(IsDBNull(CurrentRow.Item("Address")), "", CurrentRow.Item("Address"))
      Address2__String = IIf(IsDBNull(CurrentRow.Item("Address2")), "", CurrentRow.Item("Address2"))
      City__String = IIf(IsDBNull(CurrentRow.Item("City")), "", CurrentRow.Item("City"))
      State__String = IIf(IsDBNull(CurrentRow.Item("State")), "", CurrentRow.Item("State"))
      Zip__String = IIf(IsDBNull(CurrentRow.Item("Zip")), "", CurrentRow.Item("Zip"))
      HomePhone__String = IIf(IsDBNull(CurrentRow.Item("HomePhone")), "", CurrentRow.Item("HomePhone"))
      WorkPhone__String = IIf(IsDBNull(CurrentRow.Item("WorkPhone")), "", CurrentRow.Item("WorkPhone"))
      CellPhone__String = IIf(IsDBNull(CurrentRow.Item("CellPhone")), "", CurrentRow.Item("CellPhone"))
      Email__String = IIf(IsDBNull(CurrentRow.Item("Email")), "", CurrentRow.Item("Email"))
      GuestNotes__String = IIf(IsDBNull(CurrentRow.Item("GuestNotes")), "", CurrentRow.Item("GuestNotes"))
      CellPhone2__String = IIf(IsDBNull(CurrentRow.Item("CellPhone2")), "", CurrentRow.Item("CellPhone2"))
      Email2__String = IIf(IsDBNull(CurrentRow.Item("Email2")), "", CurrentRow.Item("Email2"))
      Country__String = IIf(IsDBNull(CurrentRow.Item("Country")), "", CurrentRow.Item("Country"))
      QBListID__String = IIf(IsDBNull(CurrentRow.Item("QBListID")), "", CurrentRow.Item("QBListID"))
      AuthNETCustomerProfileID__String = IIf(IsDBNull(CurrentRow.Item("AuthNETCustomerProfileID")), "", CurrentRow.Item("AuthNETCustomerProfileID"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [Guests](")
    If sInsUpdGuestName.ToString <> "" Then
      sb.Append("GuestName,")
      sbv.Append(sInsUpdGuestName & ",")
    End If
    If sInsUpdGuestFirstName.ToString <> "" Then
      sb.Append("GuestFirstName,")
      sbv.Append(sInsUpdGuestFirstName & ",")
    End If
    If sInsUpdGuestLastName.ToString <> "" Then
      sb.Append("GuestLastName,")
      sbv.Append(sInsUpdGuestLastName & ",")
    End If
    If sInsUpdAddress.ToString <> "" Then
      sb.Append("Address,")
      sbv.Append(sInsUpdAddress & ",")
    End If
    If sInsUpdAddress2.ToString <> "" Then
      sb.Append("Address2,")
      sbv.Append(sInsUpdAddress2 & ",")
    End If
    If sInsUpdCity.ToString <> "" Then
      sb.Append("City,")
      sbv.Append(sInsUpdCity & ",")
    End If
    If sInsUpdState.ToString <> "" Then
      sb.Append("State,")
      sbv.Append(sInsUpdState & ",")
    End If
    If sInsUpdZip.ToString <> "" Then
      sb.Append("Zip,")
      sbv.Append(sInsUpdZip & ",")
    End If
    If sInsUpdHomePhone.ToString <> "" Then
      sb.Append("HomePhone,")
      sbv.Append(sInsUpdHomePhone & ",")
    End If
    If sInsUpdWorkPhone.ToString <> "" Then
      sb.Append("WorkPhone,")
      sbv.Append(sInsUpdWorkPhone & ",")
    End If
    If sInsUpdCellPhone.ToString <> "" Then
      sb.Append("CellPhone,")
      sbv.Append(sInsUpdCellPhone & ",")
    End If
    If sInsUpdEmail.ToString <> "" Then
      sb.Append("Email,")
      sbv.Append(sInsUpdEmail & ",")
    End If
    If sInsUpdGuestNotes.ToString <> "" Then
      sb.Append("GuestNotes,")
      sbv.Append(sInsUpdGuestNotes & ",")
    End If
    If sInsUpdCellPhone2.ToString <> "" Then
      sb.Append("CellPhone2,")
      sbv.Append(sInsUpdCellPhone2 & ",")
    End If
    If sInsUpdEmail2.ToString <> "" Then
      sb.Append("Email2,")
      sbv.Append(sInsUpdEmail2 & ",")
    End If
    If sInsUpdCountry.ToString <> "" Then
      sb.Append("Country,")
      sbv.Append(sInsUpdCountry & ",")
    End If
    If sInsUpdQBListID.ToString <> "" Then
      sb.Append("QBListID,")
      sbv.Append(sInsUpdQBListID & ",")
    End If
    If sInsUpdAuthNETCustomerProfileID.ToString <> "" Then
      sb.Append("AuthNETCustomerProfileID,")
      sbv.Append(sInsUpdAuthNETCustomerProfileID & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(Guest_ID) from [Guests]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    Guest_ID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [Guests] Set ")
    If sInsUpdGuestName.ToString <> "" Then sb.Append("GuestName=" & sInsUpdGuestName & ",")
    If sInsUpdGuestFirstName.ToString <> "" Then sb.Append("GuestFirstName=" & sInsUpdGuestFirstName & ",")
    If sInsUpdGuestLastName.ToString <> "" Then sb.Append("GuestLastName=" & sInsUpdGuestLastName & ",")
    If sInsUpdAddress.ToString <> "" Then sb.Append("Address=" & sInsUpdAddress & ",")
    If sInsUpdAddress2.ToString <> "" Then sb.Append("Address2=" & sInsUpdAddress2 & ",")
    If sInsUpdCity.ToString <> "" Then sb.Append("City=" & sInsUpdCity & ",")
    If sInsUpdState.ToString <> "" Then sb.Append("State=" & sInsUpdState & ",")
    If sInsUpdZip.ToString <> "" Then sb.Append("Zip=" & sInsUpdZip & ",")
    If sInsUpdHomePhone.ToString <> "" Then sb.Append("HomePhone=" & sInsUpdHomePhone & ",")
    If sInsUpdWorkPhone.ToString <> "" Then sb.Append("WorkPhone=" & sInsUpdWorkPhone & ",")
    If sInsUpdCellPhone.ToString <> "" Then sb.Append("CellPhone=" & sInsUpdCellPhone & ",")
    If sInsUpdEmail.ToString <> "" Then sb.Append("Email=" & sInsUpdEmail & ",")
    If sInsUpdGuestNotes.ToString <> "" Then sb.Append("GuestNotes=" & sInsUpdGuestNotes & ",")
    If sInsUpdCellPhone2.ToString <> "" Then sb.Append("CellPhone2=" & sInsUpdCellPhone2 & ",")
    If sInsUpdEmail2.ToString <> "" Then sb.Append("Email2=" & sInsUpdEmail2 & ",")
    If sInsUpdCountry.ToString <> "" Then sb.Append("Country=" & sInsUpdCountry & ",")
    If sInsUpdQBListID.ToString <> "" Then sb.Append("QBListID=" & sInsUpdQBListID & ",")
    If sInsUpdAuthNETCustomerProfileID.ToString <> "" Then sb.Append("AuthNETCustomerProfileID=" & sInsUpdAuthNETCustomerProfileID & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where Guest_ID=" & sInsUpdGuest_ID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [Guests] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("Guest_ID=" & sInsUpdGuest_ID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TableProperties

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iProperty_ID As Int32
  Private sInsUpdProperty_ID As String
  Property Property_ID_PK__Integer() As Int32
    Get
      Return iProperty_ID
    End Get
    Set(ByVal Value As Int32)
      iProperty_ID = Value
      sInsUpdProperty_ID = oUtil.FixParam(iProperty_ID, True)
    End Set
  End Property

  Private iCategory_ID As Int32
  Private sInsUpdCategory_ID As String
  Property Category_ID__Integer() As Int32
    Get
      Return iCategory_ID
    End Get
    Set(ByVal Value As Int32)
      iCategory_ID = Value
      sInsUpdCategory_ID = oUtil.FixParam(iCategory_ID, True)
    End Set
  End Property

  Private sPropertyName As String
  Private sInsUpdPropertyName As String
  Property PropertyName__String() As String
    Get
      Return sPropertyName
    End Get
    Set(ByVal Value As String)
      sPropertyName = Value
      sInsUpdPropertyName = oUtil.FixParam(sPropertyName, True)
    End Set
  End Property

  Private sAddress As String
  Private sInsUpdAddress As String
  Property Address__String() As String
    Get
      Return sAddress
    End Get
    Set(ByVal Value As String)
      sAddress = Value
      sInsUpdAddress = oUtil.FixParam(sAddress, True)
    End Set
  End Property

  Private sAddress2 As String
  Private sInsUpdAddress2 As String
  Property Address2__String() As String
    Get
      Return sAddress2
    End Get
    Set(ByVal Value As String)
      sAddress2 = Value
      sInsUpdAddress2 = oUtil.FixParam(sAddress2, True)
    End Set
  End Property

  Private sCity As String
  Private sInsUpdCity As String
  Property City__String() As String
    Get
      Return sCity
    End Get
    Set(ByVal Value As String)
      sCity = Value
      sInsUpdCity = oUtil.FixParam(sCity, True)
    End Set
  End Property

  Private sZip As String
  Private sInsUpdZip As String
  Property Zip__String() As String
    Get
      Return sZip
    End Get
    Set(ByVal Value As String)
      sZip = Value
      sInsUpdZip = oUtil.FixParam(sZip, True)
    End Set
  End Property

  Private dMilesToCoop As Double
  Private sInsUpdMilesToCoop As String
  Property MilesToCoop__Numeric() As Double
    Get
      Return dMilesToCoop
    End Get
    Set(ByVal Value As Double)
      dMilesToCoop = Value
      sInsUpdMilesToCoop = oUtil.FixParam(dMilesToCoop, True)
    End Set
  End Property

  Private dMilesToDreams As Double
  Private sInsUpdMilesToDreams As String
  Property MilesToDreams__Numeric() As Double
    Get
      Return dMilesToDreams
    End Get
    Set(ByVal Value As Double)
      dMilesToDreams = Value
      sInsUpdMilesToDreams = oUtil.FixParam(dMilesToDreams, True)
    End Set
  End Property

  Private sPhone As String
  Private sInsUpdPhone As String
  Property Phone__String() As String
    Get
      Return sPhone
    End Get
    Set(ByVal Value As String)
      sPhone = Value
      sInsUpdPhone = oUtil.FixParam(sPhone, True)
    End Set
  End Property

  Private iSleeps As Int32
  Private sInsUpdSleeps As String
  Property Sleeps__Integer() As Int32
    Get
      Return iSleeps
    End Get
    Set(ByVal Value As Int32)
      iSleeps = Value
      sInsUpdSleeps = oUtil.FixParam(iSleeps, True)
    End Set
  End Property

  Private dBedrooms As Double
  Private sInsUpdBedrooms As String
  Property Bedrooms__Numeric() As Double
    Get
      Return dBedrooms
    End Get
    Set(ByVal Value As Double)
      dBedrooms = Value
      sInsUpdBedrooms = oUtil.FixParam(dBedrooms, True)
    End Set
  End Property

  Private dBaths As Double
  Private sInsUpdBaths As String
  Property Baths__Numeric() As Double
    Get
      Return dBaths
    End Get
    Set(ByVal Value As Double)
      dBaths = Value
      sInsUpdBaths = oUtil.FixParam(dBaths, True)
    End Set
  End Property

  Private iHost_ID As Int32
  Private sInsUpdHost_ID As String
  Property Host_ID__Integer() As Int32
    Get
      Return iHost_ID
    End Get
    Set(ByVal Value As Int32)
      iHost_ID = Value
      sInsUpdHost_ID = oUtil.FixParam(iHost_ID, True)
    End Set
  End Property

  Private dSummerRate As Double
  Private sInsUpdSummerRate As String
  Property SummerRate__Numeric() As Double
    Get
      Return dSummerRate
    End Get
    Set(ByVal Value As Double)
      dSummerRate = Value
      sInsUpdSummerRate = oUtil.FixParam(dSummerRate, True)
    End Set
  End Property

  Private dDamageDeposit As Double
  Private sInsUpdDamageDeposit As String
  Property DamageDeposit__Numeric() As Double
    Get
      Return dDamageDeposit
    End Get
    Set(ByVal Value As Double)
      dDamageDeposit = Value
      sInsUpdDamageDeposit = oUtil.FixParam(dDamageDeposit, True)
    End Set
  End Property

  Private iCableTV As Int32
  Private sInsUpdCableTV As String
  Property CableTV__Integer() As Int32
    Get
      Return iCableTV
    End Get
    Set(ByVal Value As Int32)
      iCableTV = Value
      sInsUpdCableTV = oUtil.FixParam(iCableTV, True)
    End Set
  End Property

  Private iVCR As Int32
  Private sInsUpdVCR As String
  Property VCR__Integer() As Int32
    Get
      Return iVCR
    End Get
    Set(ByVal Value As Int32)
      iVCR = Value
      sInsUpdVCR = oUtil.FixParam(iVCR, True)
    End Set
  End Property

  Private iGrill As Int32
  Private sInsUpdGrill As String
  Property Grill__Integer() As Int32
    Get
      Return iGrill
    End Get
    Set(ByVal Value As Int32)
      iGrill = Value
      sInsUpdGrill = oUtil.FixParam(iGrill, True)
    End Set
  End Property

  Private iWindowAC As Int32
  Private sInsUpdWindowAC As String
  Property WindowAC__Integer() As Int32
    Get
      Return iWindowAC
    End Get
    Set(ByVal Value As Int32)
      iWindowAC = Value
      sInsUpdWindowAC = oUtil.FixParam(iWindowAC, True)
    End Set
  End Property

  Private iWasherDryer As Int32
  Private sInsUpdWasherDryer As String
  Property WasherDryer__Integer() As Int32
    Get
      Return iWasherDryer
    End Get
    Set(ByVal Value As Int32)
      iWasherDryer = Value
      sInsUpdWasherDryer = oUtil.FixParam(iWasherDryer, True)
    End Set
  End Property

  Private iTelephone As Int32
  Private sInsUpdTelephone As String
  Property Telephone__Integer() As Int32
    Get
      Return iTelephone
    End Get
    Set(ByVal Value As Int32)
      iTelephone = Value
      sInsUpdTelephone = oUtil.FixParam(iTelephone, True)
    End Set
  End Property

  Private iBroadband As Int32
  Private sInsUpdBroadband As String
  Property Broadband__Integer() As Int32
    Get
      Return iBroadband
    End Get
    Set(ByVal Value As Int32)
      iBroadband = Value
      sInsUpdBroadband = oUtil.FixParam(iBroadband, True)
    End Set
  End Property

  Private iHandicap As Int32
  Private sInsUpdHandicap As String
  Property Handicap__Integer() As Int32
    Get
      Return iHandicap
    End Get
    Set(ByVal Value As Int32)
      iHandicap = Value
      sInsUpdHandicap = oUtil.FixParam(iHandicap, True)
    End Set
  End Property

  Private iWheelchairAccess As Int32
  Private sInsUpdWheelchairAccess As String
  Property WheelchairAccess__Integer() As Int32
    Get
      Return iWheelchairAccess
    End Get
    Set(ByVal Value As Int32)
      iWheelchairAccess = Value
      sInsUpdWheelchairAccess = oUtil.FixParam(iWheelchairAccess, True)
    End Set
  End Property

  Private sPropertyNotes As String
  Private sInsUpdPropertyNotes As String
  Property PropertyNotes__String() As String
    Get
      Return sPropertyNotes
    End Get
    Set(ByVal Value As String)
      sPropertyNotes = Value
      sInsUpdPropertyNotes = oUtil.FixParam(sPropertyNotes, True)
    End Set
  End Property

  Private iCommission_Pct As Int32
  Private sInsUpdCommission_Pct As String
  Property Commission_Pct__Integer() As Int32
    Get
      Return iCommission_Pct
    End Get
    Set(ByVal Value As Int32)
      iCommission_Pct = Value
      sInsUpdCommission_Pct = oUtil.FixParam(iCommission_Pct, True)
    End Set
  End Property

  Private iTaxRate As Int32
  Private sInsUpdTaxRate As String
  Property TaxRate__Integer() As Int32
    Get
      Return iTaxRate
    End Get
    Set(ByVal Value As Int32)
      iTaxRate = Value
      sInsUpdTaxRate = oUtil.FixParam(iTaxRate, True)
    End Set
  End Property

  Private sCollectTax As String
  Private sInsUpdCollectTax As String
  Property CollectTax__String() As String
    Get
      Return sCollectTax
    End Get
    Set(ByVal Value As String)
      sCollectTax = Value
      sInsUpdCollectTax = oUtil.FixParam(sCollectTax, True)
    End Set
  End Property

  Private sTaxIDNumber As String
  Private sInsUpdTaxIDNumber As String
  Property TaxIDNumber__String() As String
    Get
      Return sTaxIDNumber
    End Get
    Set(ByVal Value As String)
      sTaxIDNumber = Value
      sInsUpdTaxIDNumber = oUtil.FixParam(sTaxIDNumber, True)
    End Set
  End Property

  Private sStatus As String
  Private sInsUpdStatus As String
  Property Status__String() As String
    Get
      Return sStatus
    End Get
    Set(ByVal Value As String)
      sStatus = Value
      sInsUpdStatus = oUtil.FixParam(sStatus, True)
    End Set
  End Property

  Private sWebPage As String
  Private sInsUpdWebPage As String
  Property WebPage__String() As String
    Get
      Return sWebPage
    End Get
    Set(ByVal Value As String)
      sWebPage = Value
      sInsUpdWebPage = oUtil.FixParam(sWebPage, True)
    End Set
  End Property

  Private dLongitude As Double
  Private sInsUpdLongitude As String
  Property Longitude__Numeric() As Double
    Get
      Return dLongitude
    End Get
    Set(ByVal Value As Double)
      dLongitude = Value
      sInsUpdLongitude = oUtil.FixParam(dLongitude, True)
    End Set
  End Property

  Private dLatitude As Double
  Private sInsUpdLatitude As String
  Property Latitude__Numeric() As Double
    Get
      Return dLatitude
    End Get
    Set(ByVal Value As Double)
      dLatitude = Value
      sInsUpdLatitude = oUtil.FixParam(dLatitude, True)
    End Set
  End Property

  Private dDistance2CDP As Double
  Private sInsUpdDistance2CDP As String
  Property Distance2CDP__Numeric() As Double
    Get
      Return dDistance2CDP
    End Get
    Set(ByVal Value As Double)
      dDistance2CDP = Value
      sInsUpdDistance2CDP = oUtil.FixParam(dDistance2CDP, True)
    End Set
  End Property

  Private dDistance2BW As Double
  Private sInsUpdDistance2BW As String
  Property Distance2BW__Numeric() As Double
    Get
      Return dDistance2BW
    End Get
    Set(ByVal Value As Double)
      dDistance2BW = Value
      sInsUpdDistance2BW = oUtil.FixParam(dDistance2BW, True)
    End Set
  End Property

  Private dDistance2ASV As Double
  Private sInsUpdDistance2ASV As String
  Property Distance2ASV__Numeric() As Double
    Get
      Return dDistance2ASV
    End Get
    Set(ByVal Value As Double)
      dDistance2ASV = Value
      sInsUpdDistance2ASV = oUtil.FixParam(dDistance2ASV, True)
    End Set
  End Property

  Private dDistance2Coop As Double
  Private sInsUpdDistance2Coop As String
  Property Distance2Coop__Numeric() As Double
    Get
      Return dDistance2Coop
    End Get
    Set(ByVal Value As Double)
      dDistance2Coop = Value
      sInsUpdDistance2Coop = oUtil.FixParam(dDistance2Coop, True)
    End Set
  End Property

  Private iDVD As Int32
  Private sInsUpdDVD As String
  Property DVD__Integer() As Int32
    Get
      Return iDVD
    End Get
    Set(ByVal Value As Int32)
      iDVD = Value
      sInsUpdDVD = oUtil.FixParam(iDVD, True)
    End Set
  End Property

  Private iCentralAC As Int32
  Private sInsUpdCentralAC As String
  Property CentralAC__Integer() As Int32
    Get
      Return iCentralAC
    End Get
    Set(ByVal Value As Int32)
      iCentralAC = Value
      sInsUpdCentralAC = oUtil.FixParam(iCentralAC, True)
    End Set
  End Property

  Private iDishSatellite As Int32
  Private sInsUpdDishSatellite As String
  Property DishSatellite__Integer() As Int32
    Get
      Return iDishSatellite
    End Get
    Set(ByVal Value As Int32)
      iDishSatellite = Value
      sInsUpdDishSatellite = oUtil.FixParam(iDishSatellite, True)
    End Set
  End Property

  Private iDialup As Int32
  Private sInsUpdDialup As String
  Property Dialup__Integer() As Int32
    Get
      Return iDialup
    End Get
    Set(ByVal Value As Int32)
      iDialup = Value
      sInsUpdDialup = oUtil.FixParam(iDialup, True)
    End Set
  End Property

  Private iSquareFootage As Int32
  Private sInsUpdSquareFootage As String
  Property SquareFootage__Integer() As Int32
    Get
      Return iSquareFootage
    End Get
    Set(ByVal Value As Int32)
      iSquareFootage = Value
      sInsUpdSquareFootage = oUtil.FixParam(iSquareFootage, True)
    End Set
  End Property

  Private iKingBeds As Int32
  Private sInsUpdKingBeds As String
  Property KingBeds__Integer() As Int32
    Get
      Return iKingBeds
    End Get
    Set(ByVal Value As Int32)
      iKingBeds = Value
      sInsUpdKingBeds = oUtil.FixParam(iKingBeds, True)
    End Set
  End Property

  Private iQueenBeds As Int32
  Private sInsUpdQueenBeds As String
  Property QueenBeds__Integer() As Int32
    Get
      Return iQueenBeds
    End Get
    Set(ByVal Value As Int32)
      iQueenBeds = Value
      sInsUpdQueenBeds = oUtil.FixParam(iQueenBeds, True)
    End Set
  End Property

  Private iDoubleBeds As Int32
  Private sInsUpdDoubleBeds As String
  Property DoubleBeds__Integer() As Int32
    Get
      Return iDoubleBeds
    End Get
    Set(ByVal Value As Int32)
      iDoubleBeds = Value
      sInsUpdDoubleBeds = oUtil.FixParam(iDoubleBeds, True)
    End Set
  End Property

  Private iTwinBeds As Int32
  Private sInsUpdTwinBeds As String
  Property TwinBeds__Integer() As Int32
    Get
      Return iTwinBeds
    End Get
    Set(ByVal Value As Int32)
      iTwinBeds = Value
      sInsUpdTwinBeds = oUtil.FixParam(iTwinBeds, True)
    End Set
  End Property

  Private iBunkBeds As Int32
  Private sInsUpdBunkBeds As String
  Property BunkBeds__Integer() As Int32
    Get
      Return iBunkBeds
    End Get
    Set(ByVal Value As Int32)
      iBunkBeds = Value
      sInsUpdBunkBeds = oUtil.FixParam(iBunkBeds, True)
    End Set
  End Property

  Private iSleeperSofa As Int32
  Private sInsUpdSleeperSofa As String
  Property SleeperSofa__Integer() As Int32
    Get
      Return iSleeperSofa
    End Get
    Set(ByVal Value As Int32)
      iSleeperSofa = Value
      sInsUpdSleeperSofa = oUtil.FixParam(iSleeperSofa, True)
    End Set
  End Property

  Private iFuton As Int32
  Private sInsUpdFuton As String
  Property Futon__Integer() As Int32
    Get
      Return iFuton
    End Get
    Set(ByVal Value As Int32)
      iFuton = Value
      sInsUpdFuton = oUtil.FixParam(iFuton, True)
    End Set
  End Property

  Private iDishwasher As Int32
  Private sInsUpdDishwasher As String
  Property Dishwasher__Integer() As Int32
    Get
      Return iDishwasher
    End Get
    Set(ByVal Value As Int32)
      iDishwasher = Value
      sInsUpdDishwasher = oUtil.FixParam(iDishwasher, True)
    End Set
  End Property

  Private iPool As Int32
  Private sInsUpdPool As String
  Property Pool__Integer() As Int32
    Get
      Return iPool
    End Get
    Set(ByVal Value As Int32)
      iPool = Value
      sInsUpdPool = oUtil.FixParam(iPool, True)
    End Set
  End Property

  Private iWireless As Int32
  Private sInsUpdWireless As String
  Property Wireless__Integer() As Int32
    Get
      Return iWireless
    End Get
    Set(ByVal Value As Int32)
      iWireless = Value
      sInsUpdWireless = oUtil.FixParam(iWireless, True)
    End Set
  End Property

  Private iPrivatePond As Int32
  Private sInsUpdPrivatePond As String
  Property PrivatePond__Integer() As Int32
    Get
      Return iPrivatePond
    End Get
    Set(ByVal Value As Int32)
      iPrivatePond = Value
      sInsUpdPrivatePond = oUtil.FixParam(iPrivatePond, True)
    End Set
  End Property

  Private iTeamParties As Int32
  Private sInsUpdTeamParties As String
  Property TeamParties__Integer() As Int32
    Get
      Return iTeamParties
    End Get
    Set(ByVal Value As Int32)
      iTeamParties = Value
      sInsUpdTeamParties = oUtil.FixParam(iTeamParties, True)
    End Set
  End Property

  Private sWebImage As String
  Private sInsUpdWebImage As String
  Property WebImage__String() As String
    Get
      Return sWebImage
    End Get
    Set(ByVal Value As String)
      sWebImage = Value
      sInsUpdWebImage = oUtil.FixParam(sWebImage, True)
    End Set
  End Property

  Private iPropertyGroup As Int32
  Private sInsUpdPropertyGroup As String
  Property PropertyGroup__Integer() As Int32
    Get
      Return iPropertyGroup
    End Get
    Set(ByVal Value As Int32)
      iPropertyGroup = Value
      sInsUpdPropertyGroup = oUtil.FixParam(iPropertyGroup, True)
    End Set
  End Property

  Private sQBListID As String
  Private sInsUpdQBListID As String
  Property QBListID__String() As String
    Get
      Return sQBListID
    End Get
    Set(ByVal Value As String)
      sQBListID = Value
      sInsUpdQBListID = oUtil.FixParam(sQBListID, True)
    End Set
  End Property

  Private sCheckName As String
  Private sInsUpdCheckName As String
  Property CheckName__String() As String
    Get
      Return sCheckName
    End Get
    Set(ByVal Value As String)
      sCheckName = Value
      sInsUpdCheckName = oUtil.FixParam(sCheckName, True)
    End Set
  End Property

  Private dDiscountedRate As Double
  Private sInsUpdDiscountedRate As String
  Property DiscountedRate__Numeric() As Double
    Get
      Return dDiscountedRate
    End Get
    Set(ByVal Value As Double)
      dDiscountedRate = Value
      sInsUpdDiscountedRate = oUtil.FixParam(dDiscountedRate, True)
    End Set
  End Property

  Private sWebMasterDescription As String
  Private sInsUpdWebMasterDescription As String
  Property WebMasterDescription__String() As String
    Get
      Return sWebMasterDescription
    End Get
    Set(ByVal Value As String)
      sWebMasterDescription = Value
      sInsUpdWebMasterDescription = oUtil.FixParam(sWebMasterDescription, True)
    End Set
  End Property

  Private sWebDetailsDescription As String
  Private sInsUpdWebDetailsDescription As String
  Property WebDetailsDescription__String() As String
    Get
      Return sWebDetailsDescription
    End Get
    Set(ByVal Value As String)
      sWebDetailsDescription = Value
      sInsUpdWebDetailsDescription = oUtil.FixParam(sWebDetailsDescription, True)
    End Set
  End Property

  Private sWebMasterTitle As String
  Private sInsUpdWebMasterTitle As String
  Property WebMasterTitle__String() As String
    Get
      Return sWebMasterTitle
    End Get
    Set(ByVal Value As String)
      sWebMasterTitle = Value
      sInsUpdWebMasterTitle = oUtil.FixParam(sWebMasterTitle, True)
    End Set
  End Property

  Private sWebDetailsTitle As String
  Private sInsUpdWebDetailsTitle As String
  Property WebDetailsTitle__String() As String
    Get
      Return sWebDetailsTitle
    End Get
    Set(ByVal Value As String)
      sWebDetailsTitle = Value
      sInsUpdWebDetailsTitle = oUtil.FixParam(sWebDetailsTitle, True)
    End Set
  End Property

  Private sWebDetailsLeftSection As String
  Private sInsUpdWebDetailsLeftSection As String
  Property WebDetailsLeftSection__String() As String
    Get
      Return sWebDetailsLeftSection
    End Get
    Set(ByVal Value As String)
      sWebDetailsLeftSection = Value
      sInsUpdWebDetailsLeftSection = oUtil.FixParam(sWebDetailsLeftSection, True)
    End Set
  End Property

  Private sWebDetailsRightSection As String
  Private sInsUpdWebDetailsRightSection As String
  Property WebDetailsRightSection__String() As String
    Get
      Return sWebDetailsRightSection
    End Get
    Set(ByVal Value As String)
      sWebDetailsRightSection = Value
      sInsUpdWebDetailsRightSection = oUtil.FixParam(sWebDetailsRightSection, True)
    End Set
  End Property

  Private iCategory_ID2 As Int32
  Private sInsUpdCategory_ID2 As String
  Property Category_ID2__Integer() As Int32
    Get
      Return iCategory_ID2
    End Get
    Set(ByVal Value As Int32)
      iCategory_ID2 = Value
      sInsUpdCategory_ID2 = oUtil.FixParam(iCategory_ID2, True)
    End Set
  End Property

  Public Sub Clear()
    iProperty_ID = 0
    sInsUpdProperty_ID = ""
    iCategory_ID = 0
    sInsUpdCategory_ID = ""
    sPropertyName = ""
    sInsUpdPropertyName = ""
    sAddress = ""
    sInsUpdAddress = ""
    sAddress2 = ""
    sInsUpdAddress2 = ""
    sCity = ""
    sInsUpdCity = ""
    sZip = ""
    sInsUpdZip = ""
    dMilesToCoop = 0.0
    sInsUpdMilesToCoop = ""
    dMilesToDreams = 0.0
    sInsUpdMilesToDreams = ""
    sPhone = ""
    sInsUpdPhone = ""
    iSleeps = 0
    sInsUpdSleeps = ""
    dBedrooms = 0.0
    sInsUpdBedrooms = ""
    dBaths = 0.0
    sInsUpdBaths = ""
    iHost_ID = 0
    sInsUpdHost_ID = ""
    dSummerRate = 0.0
    sInsUpdSummerRate = ""
    dDamageDeposit = 0.0
    sInsUpdDamageDeposit = ""
    iCableTV = 0
    sInsUpdCableTV = ""
    iVCR = 0
    sInsUpdVCR = ""
    iGrill = 0
    sInsUpdGrill = ""
    iWindowAC = 0
    sInsUpdWindowAC = ""
    iWasherDryer = 0
    sInsUpdWasherDryer = ""
    iTelephone = 0
    sInsUpdTelephone = ""
    iBroadband = 0
    sInsUpdBroadband = ""
    iHandicap = 0
    sInsUpdHandicap = ""
    iWheelchairAccess = 0
    sInsUpdWheelchairAccess = ""
    sPropertyNotes = ""
    sInsUpdPropertyNotes = ""
    iCommission_Pct = 0
    sInsUpdCommission_Pct = ""
    iTaxRate = 0
    sInsUpdTaxRate = ""
    sCollectTax = ""
    sInsUpdCollectTax = ""
    sTaxIDNumber = ""
    sInsUpdTaxIDNumber = ""
    sStatus = ""
    sInsUpdStatus = ""
    sWebPage = ""
    sInsUpdWebPage = ""
    dLongitude = 0.0
    sInsUpdLongitude = ""
    dLatitude = 0.0
    sInsUpdLatitude = ""
    dDistance2CDP = 0.0
    sInsUpdDistance2CDP = ""
    dDistance2BW = 0.0
    sInsUpdDistance2BW = ""
    dDistance2ASV = 0.0
    sInsUpdDistance2ASV = ""
    dDistance2Coop = 0.0
    sInsUpdDistance2Coop = ""
    iDVD = 0
    sInsUpdDVD = ""
    iCentralAC = 0
    sInsUpdCentralAC = ""
    iDishSatellite = 0
    sInsUpdDishSatellite = ""
    iDialup = 0
    sInsUpdDialup = ""
    iSquareFootage = 0
    sInsUpdSquareFootage = ""
    iKingBeds = 0
    sInsUpdKingBeds = ""
    iQueenBeds = 0
    sInsUpdQueenBeds = ""
    iDoubleBeds = 0
    sInsUpdDoubleBeds = ""
    iTwinBeds = 0
    sInsUpdTwinBeds = ""
    iBunkBeds = 0
    sInsUpdBunkBeds = ""
    iSleeperSofa = 0
    sInsUpdSleeperSofa = ""
    iFuton = 0
    sInsUpdFuton = ""
    iDishwasher = 0
    sInsUpdDishwasher = ""
    iPool = 0
    sInsUpdPool = ""
    iWireless = 0
    sInsUpdWireless = ""
    iPrivatePond = 0
    sInsUpdPrivatePond = ""
    iTeamParties = 0
    sInsUpdTeamParties = ""
    sWebImage = ""
    sInsUpdWebImage = ""
    iPropertyGroup = 0
    sInsUpdPropertyGroup = ""
    sQBListID = ""
    sInsUpdQBListID = ""
    sCheckName = ""
    sInsUpdCheckName = ""
    dDiscountedRate = 0.0
    sInsUpdDiscountedRate = ""
    sWebMasterDescription = ""
    sInsUpdWebMasterDescription = ""
    sWebDetailsDescription = ""
    sInsUpdWebDetailsDescription = ""
    sWebMasterTitle = ""
    sInsUpdWebMasterTitle = ""
    sWebDetailsTitle = ""
    sInsUpdWebDetailsTitle = ""
    sWebDetailsLeftSection = ""
    sInsUpdWebDetailsLeftSection = ""
    sWebDetailsRightSection = ""
    sInsUpdWebDetailsRightSection = ""
    iCategory_ID2 = 0
    sInsUpdCategory_ID2 = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdProperty_ID.ToString = "'-12345'" Then sb.Append("Property_ID,")
        If sInsUpdCategory_ID.ToString = "'-12345'" Then sb.Append("Category_ID,")
        If sInsUpdPropertyName.ToString = "'Select'" Then sb.Append("PropertyName,")
        If sInsUpdAddress.ToString = "'Select'" Then sb.Append("Address,")
        If sInsUpdAddress2.ToString = "'Select'" Then sb.Append("Address2,")
        If sInsUpdCity.ToString = "'Select'" Then sb.Append("City,")
        If sInsUpdZip.ToString = "'Select'" Then sb.Append("Zip,")
        If sInsUpdMilesToCoop.ToString = "'-12345'" Then sb.Append("MilesToCoop,")
        If sInsUpdMilesToDreams.ToString = "'-12345'" Then sb.Append("MilesToDreams,")
        If sInsUpdPhone.ToString = "'Select'" Then sb.Append("Phone,")
        If sInsUpdSleeps.ToString = "'-12345'" Then sb.Append("Sleeps,")
        If sInsUpdBedrooms.ToString = "'-12345'" Then sb.Append("Bedrooms,")
        If sInsUpdBaths.ToString = "'-12345'" Then sb.Append("Baths,")
        If sInsUpdHost_ID.ToString = "'-12345'" Then sb.Append("Host_ID,")
        If sInsUpdSummerRate.ToString = "'-12345'" Then sb.Append("SummerRate,")
        If sInsUpdDamageDeposit.ToString = "'-12345'" Then sb.Append("DamageDeposit,")
        If sInsUpdCableTV.ToString = "'-12345'" Then sb.Append("CableTV,")
        If sInsUpdVCR.ToString = "'-12345'" Then sb.Append("VCR,")
        If sInsUpdGrill.ToString = "'-12345'" Then sb.Append("Grill,")
        If sInsUpdWindowAC.ToString = "'-12345'" Then sb.Append("WindowAC,")
        If sInsUpdWasherDryer.ToString = "'-12345'" Then sb.Append("WasherDryer,")
        If sInsUpdTelephone.ToString = "'-12345'" Then sb.Append("Telephone,")
        If sInsUpdBroadband.ToString = "'-12345'" Then sb.Append("Broadband,")
        If sInsUpdHandicap.ToString = "'-12345'" Then sb.Append("Handicap,")
        If sInsUpdWheelchairAccess.ToString = "'-12345'" Then sb.Append("WheelchairAccess,")
        If sInsUpdPropertyNotes.ToString = "'Select'" Then sb.Append("PropertyNotes,")
        If sInsUpdCommission_Pct.ToString = "'-12345'" Then sb.Append("Commission_Pct,")
        If sInsUpdTaxRate.ToString = "'-12345'" Then sb.Append("TaxRate,")
        If sInsUpdCollectTax.ToString = "'Select'" Then sb.Append("CollectTax,")
        If sInsUpdTaxIDNumber.ToString = "'Select'" Then sb.Append("TaxIDNumber,")
        If sInsUpdStatus.ToString = "'Select'" Then sb.Append("Status,")
        If sInsUpdWebPage.ToString = "'Select'" Then sb.Append("WebPage,")
        If sInsUpdLongitude.ToString = "'-12345'" Then sb.Append("Longitude,")
        If sInsUpdLatitude.ToString = "'-12345'" Then sb.Append("Latitude,")
        If sInsUpdDistance2CDP.ToString = "'-12345'" Then sb.Append("Distance2CDP,")
        If sInsUpdDistance2BW.ToString = "'-12345'" Then sb.Append("Distance2BW,")
        If sInsUpdDistance2ASV.ToString = "'-12345'" Then sb.Append("Distance2ASV,")
        If sInsUpdDistance2Coop.ToString = "'-12345'" Then sb.Append("Distance2Coop,")
        If sInsUpdDVD.ToString = "'-12345'" Then sb.Append("DVD,")
        If sInsUpdCentralAC.ToString = "'-12345'" Then sb.Append("CentralAC,")
        If sInsUpdDishSatellite.ToString = "'-12345'" Then sb.Append("DishSatellite,")
        If sInsUpdDialup.ToString = "'-12345'" Then sb.Append("Dialup,")
        If sInsUpdSquareFootage.ToString = "'-12345'" Then sb.Append("SquareFootage,")
        If sInsUpdKingBeds.ToString = "'-12345'" Then sb.Append("KingBeds,")
        If sInsUpdQueenBeds.ToString = "'-12345'" Then sb.Append("QueenBeds,")
        If sInsUpdDoubleBeds.ToString = "'-12345'" Then sb.Append("DoubleBeds,")
        If sInsUpdTwinBeds.ToString = "'-12345'" Then sb.Append("TwinBeds,")
        If sInsUpdBunkBeds.ToString = "'-12345'" Then sb.Append("BunkBeds,")
        If sInsUpdSleeperSofa.ToString = "'-12345'" Then sb.Append("SleeperSofa,")
        If sInsUpdFuton.ToString = "'-12345'" Then sb.Append("Futon,")
        If sInsUpdDishwasher.ToString = "'-12345'" Then sb.Append("Dishwasher,")
        If sInsUpdPool.ToString = "'-12345'" Then sb.Append("Pool,")
        If sInsUpdWireless.ToString = "'-12345'" Then sb.Append("Wireless,")
        If sInsUpdPrivatePond.ToString = "'-12345'" Then sb.Append("PrivatePond,")
        If sInsUpdTeamParties.ToString = "'-12345'" Then sb.Append("TeamParties,")
        If sInsUpdWebImage.ToString = "'Select'" Then sb.Append("WebImage,")
        If sInsUpdPropertyGroup.ToString = "'-12345'" Then sb.Append("PropertyGroup,")
        If sInsUpdQBListID.ToString = "'Select'" Then sb.Append("QBListID,")
        If sInsUpdCheckName.ToString = "'Select'" Then sb.Append("CheckName,")
        If sInsUpdDiscountedRate.ToString = "'-12345'" Then sb.Append("DiscountedRate,")
        If sInsUpdWebMasterDescription.ToString = "'Select'" Then sb.Append("WebMasterDescription,")
        If sInsUpdWebDetailsDescription.ToString = "'Select'" Then sb.Append("WebDetailsDescription,")
        If sInsUpdWebMasterTitle.ToString = "'Select'" Then sb.Append("WebMasterTitle,")
        If sInsUpdWebDetailsTitle.ToString = "'Select'" Then sb.Append("WebDetailsTitle,")
        If sInsUpdWebDetailsLeftSection.ToString = "'Select'" Then sb.Append("WebDetailsLeftSection,")
        If sInsUpdWebDetailsRightSection.ToString = "'Select'" Then sb.Append("WebDetailsRightSection,")
        If sInsUpdCategory_ID2.ToString = "'-12345'" Then sb.Append("Category_ID2,")
      Else
        sb.Append("Property_ID,")
        sb.Append("Category_ID,")
        sb.Append("PropertyName,")
        sb.Append("Address,")
        sb.Append("Address2,")
        sb.Append("City,")
        sb.Append("Zip,")
        sb.Append("MilesToCoop,")
        sb.Append("MilesToDreams,")
        sb.Append("Phone,")
        sb.Append("Sleeps,")
        sb.Append("Bedrooms,")
        sb.Append("Baths,")
        sb.Append("Host_ID,")
        sb.Append("SummerRate,")
        sb.Append("DamageDeposit,")
        sb.Append("CableTV,")
        sb.Append("VCR,")
        sb.Append("Grill,")
        sb.Append("WindowAC,")
        sb.Append("WasherDryer,")
        sb.Append("Telephone,")
        sb.Append("Broadband,")
        sb.Append("Handicap,")
        sb.Append("WheelchairAccess,")
        sb.Append("PropertyNotes,")
        sb.Append("Commission_Pct,")
        sb.Append("TaxRate,")
        sb.Append("CollectTax,")
        sb.Append("TaxIDNumber,")
        sb.Append("Status,")
        sb.Append("WebPage,")
        sb.Append("Longitude,")
        sb.Append("Latitude,")
        sb.Append("Distance2CDP,")
        sb.Append("Distance2BW,")
        sb.Append("Distance2ASV,")
        sb.Append("Distance2Coop,")
        sb.Append("DVD,")
        sb.Append("CentralAC,")
        sb.Append("DishSatellite,")
        sb.Append("Dialup,")
        sb.Append("SquareFootage,")
        sb.Append("KingBeds,")
        sb.Append("QueenBeds,")
        sb.Append("DoubleBeds,")
        sb.Append("TwinBeds,")
        sb.Append("BunkBeds,")
        sb.Append("SleeperSofa,")
        sb.Append("Futon,")
        sb.Append("Dishwasher,")
        sb.Append("Pool,")
        sb.Append("Wireless,")
        sb.Append("PrivatePond,")
        sb.Append("TeamParties,")
        sb.Append("WebImage,")
        sb.Append("PropertyGroup,")
        sb.Append("QBListID,")
        sb.Append("CheckName,")
        sb.Append("DiscountedRate,")
        sb.Append("WebMasterDescription,")
        sb.Append("WebDetailsDescription,")
        sb.Append("WebMasterTitle,")
        sb.Append("WebDetailsTitle,")
        sb.Append("WebDetailsLeftSection,")
        sb.Append("WebDetailsRightSection,")
        sb.Append("Category_ID2,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [Properties]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdProperty_ID.ToString <> "") And (sInsUpdProperty_ID <> "'-12345'") Then sbw.Append("Property_ID=" & sInsUpdProperty_ID & " and ")
      If (sInsUpdCategory_ID.ToString <> "") And (sInsUpdCategory_ID <> "'-12345'") Then sbw.Append("Category_ID=" & sInsUpdCategory_ID & " and ")
      If (sInsUpdPropertyName.ToString <> "") And (sInsUpdPropertyName <> "'Select'") Then sbw.Append("PropertyName=" & sInsUpdPropertyName & " and ")
      If (sInsUpdAddress.ToString <> "") And (sInsUpdAddress <> "'Select'") Then sbw.Append("Address=" & sInsUpdAddress & " and ")
      If (sInsUpdAddress2.ToString <> "") And (sInsUpdAddress2 <> "'Select'") Then sbw.Append("Address2=" & sInsUpdAddress2 & " and ")
      If (sInsUpdCity.ToString <> "") And (sInsUpdCity <> "'Select'") Then sbw.Append("City=" & sInsUpdCity & " and ")
      If (sInsUpdZip.ToString <> "") And (sInsUpdZip <> "'Select'") Then sbw.Append("Zip=" & sInsUpdZip & " and ")
      If (sInsUpdMilesToCoop.ToString <> "") And (sInsUpdMilesToCoop <> "'-12345'") Then sbw.Append("MilesToCoop=" & sInsUpdMilesToCoop & " and ")
      If (sInsUpdMilesToDreams.ToString <> "") And (sInsUpdMilesToDreams <> "'-12345'") Then sbw.Append("MilesToDreams=" & sInsUpdMilesToDreams & " and ")
      If (sInsUpdPhone.ToString <> "") And (sInsUpdPhone <> "'Select'") Then sbw.Append("Phone=" & sInsUpdPhone & " and ")
      If (sInsUpdSleeps.ToString <> "") And (sInsUpdSleeps <> "'-12345'") Then sbw.Append("Sleeps=" & sInsUpdSleeps & " and ")
      If (sInsUpdBedrooms.ToString <> "") And (sInsUpdBedrooms <> "'-12345'") Then sbw.Append("Bedrooms=" & sInsUpdBedrooms & " and ")
      If (sInsUpdBaths.ToString <> "") And (sInsUpdBaths <> "'-12345'") Then sbw.Append("Baths=" & sInsUpdBaths & " and ")
      If (sInsUpdHost_ID.ToString <> "") And (sInsUpdHost_ID <> "'-12345'") Then sbw.Append("Host_ID=" & sInsUpdHost_ID & " and ")
      If (sInsUpdSummerRate.ToString <> "") And (sInsUpdSummerRate <> "'-12345'") Then sbw.Append("SummerRate=" & sInsUpdSummerRate & " and ")
      If (sInsUpdDamageDeposit.ToString <> "") And (sInsUpdDamageDeposit <> "'-12345'") Then sbw.Append("DamageDeposit=" & sInsUpdDamageDeposit & " and ")
      If (sInsUpdCableTV.ToString <> "") And (sInsUpdCableTV <> "'-12345'") Then sbw.Append("CableTV=" & sInsUpdCableTV & " and ")
      If (sInsUpdVCR.ToString <> "") And (sInsUpdVCR <> "'-12345'") Then sbw.Append("VCR=" & sInsUpdVCR & " and ")
      If (sInsUpdGrill.ToString <> "") And (sInsUpdGrill <> "'-12345'") Then sbw.Append("Grill=" & sInsUpdGrill & " and ")
      If (sInsUpdWindowAC.ToString <> "") And (sInsUpdWindowAC <> "'-12345'") Then sbw.Append("WindowAC=" & sInsUpdWindowAC & " and ")
      If (sInsUpdWasherDryer.ToString <> "") And (sInsUpdWasherDryer <> "'-12345'") Then sbw.Append("WasherDryer=" & sInsUpdWasherDryer & " and ")
      If (sInsUpdTelephone.ToString <> "") And (sInsUpdTelephone <> "'-12345'") Then sbw.Append("Telephone=" & sInsUpdTelephone & " and ")
      If (sInsUpdBroadband.ToString <> "") And (sInsUpdBroadband <> "'-12345'") Then sbw.Append("Broadband=" & sInsUpdBroadband & " and ")
      If (sInsUpdHandicap.ToString <> "") And (sInsUpdHandicap <> "'-12345'") Then sbw.Append("Handicap=" & sInsUpdHandicap & " and ")
      If (sInsUpdWheelchairAccess.ToString <> "") And (sInsUpdWheelchairAccess <> "'-12345'") Then sbw.Append("WheelchairAccess=" & sInsUpdWheelchairAccess & " and ")
      If (sInsUpdPropertyNotes.ToString <> "") And (sInsUpdPropertyNotes <> "'Select'") Then sbw.Append("PropertyNotes=" & sInsUpdPropertyNotes & " and ")
      If (sInsUpdCommission_Pct.ToString <> "") And (sInsUpdCommission_Pct <> "'-12345'") Then sbw.Append("Commission_Pct=" & sInsUpdCommission_Pct & " and ")
      If (sInsUpdTaxRate.ToString <> "") And (sInsUpdTaxRate <> "'-12345'") Then sbw.Append("TaxRate=" & sInsUpdTaxRate & " and ")
      If (sInsUpdCollectTax.ToString <> "") And (sInsUpdCollectTax <> "'Select'") Then sbw.Append("CollectTax=" & sInsUpdCollectTax & " and ")
      If (sInsUpdTaxIDNumber.ToString <> "") And (sInsUpdTaxIDNumber <> "'Select'") Then sbw.Append("TaxIDNumber=" & sInsUpdTaxIDNumber & " and ")
      If (sInsUpdStatus.ToString <> "") And (sInsUpdStatus <> "'Select'") Then sbw.Append("Status=" & sInsUpdStatus & " and ")
      If (sInsUpdWebPage.ToString <> "") And (sInsUpdWebPage <> "'Select'") Then sbw.Append("WebPage=" & sInsUpdWebPage & " and ")
      If (sInsUpdLongitude.ToString <> "") And (sInsUpdLongitude <> "'-12345'") Then sbw.Append("Longitude=" & sInsUpdLongitude & " and ")
      If (sInsUpdLatitude.ToString <> "") And (sInsUpdLatitude <> "'-12345'") Then sbw.Append("Latitude=" & sInsUpdLatitude & " and ")
      If (sInsUpdDistance2CDP.ToString <> "") And (sInsUpdDistance2CDP <> "'-12345'") Then sbw.Append("Distance2CDP=" & sInsUpdDistance2CDP & " and ")
      If (sInsUpdDistance2BW.ToString <> "") And (sInsUpdDistance2BW <> "'-12345'") Then sbw.Append("Distance2BW=" & sInsUpdDistance2BW & " and ")
      If (sInsUpdDistance2ASV.ToString <> "") And (sInsUpdDistance2ASV <> "'-12345'") Then sbw.Append("Distance2ASV=" & sInsUpdDistance2ASV & " and ")
      If (sInsUpdDistance2Coop.ToString <> "") And (sInsUpdDistance2Coop <> "'-12345'") Then sbw.Append("Distance2Coop=" & sInsUpdDistance2Coop & " and ")
      If (sInsUpdDVD.ToString <> "") And (sInsUpdDVD <> "'-12345'") Then sbw.Append("DVD=" & sInsUpdDVD & " and ")
      If (sInsUpdCentralAC.ToString <> "") And (sInsUpdCentralAC <> "'-12345'") Then sbw.Append("CentralAC=" & sInsUpdCentralAC & " and ")
      If (sInsUpdDishSatellite.ToString <> "") And (sInsUpdDishSatellite <> "'-12345'") Then sbw.Append("DishSatellite=" & sInsUpdDishSatellite & " and ")
      If (sInsUpdDialup.ToString <> "") And (sInsUpdDialup <> "'-12345'") Then sbw.Append("Dialup=" & sInsUpdDialup & " and ")
      If (sInsUpdSquareFootage.ToString <> "") And (sInsUpdSquareFootage <> "'-12345'") Then sbw.Append("SquareFootage=" & sInsUpdSquareFootage & " and ")
      If (sInsUpdKingBeds.ToString <> "") And (sInsUpdKingBeds <> "'-12345'") Then sbw.Append("KingBeds=" & sInsUpdKingBeds & " and ")
      If (sInsUpdQueenBeds.ToString <> "") And (sInsUpdQueenBeds <> "'-12345'") Then sbw.Append("QueenBeds=" & sInsUpdQueenBeds & " and ")
      If (sInsUpdDoubleBeds.ToString <> "") And (sInsUpdDoubleBeds <> "'-12345'") Then sbw.Append("DoubleBeds=" & sInsUpdDoubleBeds & " and ")
      If (sInsUpdTwinBeds.ToString <> "") And (sInsUpdTwinBeds <> "'-12345'") Then sbw.Append("TwinBeds=" & sInsUpdTwinBeds & " and ")
      If (sInsUpdBunkBeds.ToString <> "") And (sInsUpdBunkBeds <> "'-12345'") Then sbw.Append("BunkBeds=" & sInsUpdBunkBeds & " and ")
      If (sInsUpdSleeperSofa.ToString <> "") And (sInsUpdSleeperSofa <> "'-12345'") Then sbw.Append("SleeperSofa=" & sInsUpdSleeperSofa & " and ")
      If (sInsUpdFuton.ToString <> "") And (sInsUpdFuton <> "'-12345'") Then sbw.Append("Futon=" & sInsUpdFuton & " and ")
      If (sInsUpdDishwasher.ToString <> "") And (sInsUpdDishwasher <> "'-12345'") Then sbw.Append("Dishwasher=" & sInsUpdDishwasher & " and ")
      If (sInsUpdPool.ToString <> "") And (sInsUpdPool <> "'-12345'") Then sbw.Append("Pool=" & sInsUpdPool & " and ")
      If (sInsUpdWireless.ToString <> "") And (sInsUpdWireless <> "'-12345'") Then sbw.Append("Wireless=" & sInsUpdWireless & " and ")
      If (sInsUpdPrivatePond.ToString <> "") And (sInsUpdPrivatePond <> "'-12345'") Then sbw.Append("PrivatePond=" & sInsUpdPrivatePond & " and ")
      If (sInsUpdTeamParties.ToString <> "") And (sInsUpdTeamParties <> "'-12345'") Then sbw.Append("TeamParties=" & sInsUpdTeamParties & " and ")
      If (sInsUpdWebImage.ToString <> "") And (sInsUpdWebImage <> "'Select'") Then sbw.Append("WebImage=" & sInsUpdWebImage & " and ")
      If (sInsUpdPropertyGroup.ToString <> "") And (sInsUpdPropertyGroup <> "'-12345'") Then sbw.Append("PropertyGroup=" & sInsUpdPropertyGroup & " and ")
      If (sInsUpdQBListID.ToString <> "") And (sInsUpdQBListID <> "'Select'") Then sbw.Append("QBListID=" & sInsUpdQBListID & " and ")
      If (sInsUpdCheckName.ToString <> "") And (sInsUpdCheckName <> "'Select'") Then sbw.Append("CheckName=" & sInsUpdCheckName & " and ")
      If (sInsUpdDiscountedRate.ToString <> "") And (sInsUpdDiscountedRate <> "'-12345'") Then sbw.Append("DiscountedRate=" & sInsUpdDiscountedRate & " and ")
      If (sInsUpdWebMasterDescription.ToString <> "") And (sInsUpdWebMasterDescription <> "'Select'") Then sbw.Append("WebMasterDescription=" & sInsUpdWebMasterDescription & " and ")
      If (sInsUpdWebDetailsDescription.ToString <> "") And (sInsUpdWebDetailsDescription <> "'Select'") Then sbw.Append("WebDetailsDescription=" & sInsUpdWebDetailsDescription & " and ")
      If (sInsUpdWebMasterTitle.ToString <> "") And (sInsUpdWebMasterTitle <> "'Select'") Then sbw.Append("WebMasterTitle=" & sInsUpdWebMasterTitle & " and ")
      If (sInsUpdWebDetailsTitle.ToString <> "") And (sInsUpdWebDetailsTitle <> "'Select'") Then sbw.Append("WebDetailsTitle=" & sInsUpdWebDetailsTitle & " and ")
      If (sInsUpdWebDetailsLeftSection.ToString <> "") And (sInsUpdWebDetailsLeftSection <> "'Select'") Then sbw.Append("WebDetailsLeftSection=" & sInsUpdWebDetailsLeftSection & " and ")
      If (sInsUpdWebDetailsRightSection.ToString <> "") And (sInsUpdWebDetailsRightSection <> "'Select'") Then sbw.Append("WebDetailsRightSection=" & sInsUpdWebDetailsRightSection & " and ")
      If (sInsUpdCategory_ID2.ToString <> "") And (sInsUpdCategory_ID2 <> "'-12345'") Then sbw.Append("Category_ID2=" & sInsUpdCategory_ID2 & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      Property_ID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("Property_ID")), 0, CurrentRow.Item("Property_ID").ToString)
      Category_ID__Integer = IIf(IsDBNull(CurrentRow.Item("Category_ID")), 0, CurrentRow.Item("Category_ID"))
      PropertyName__String = IIf(IsDBNull(CurrentRow.Item("PropertyName")), "", CurrentRow.Item("PropertyName"))
      Address__String = IIf(IsDBNull(CurrentRow.Item("Address")), "", CurrentRow.Item("Address"))
      Address2__String = IIf(IsDBNull(CurrentRow.Item("Address2")), "", CurrentRow.Item("Address2"))
      City__String = IIf(IsDBNull(CurrentRow.Item("City")), "", CurrentRow.Item("City"))
      Zip__String = IIf(IsDBNull(CurrentRow.Item("Zip")), "", CurrentRow.Item("Zip"))
      MilesToCoop__Numeric = IIf(IsDBNull(CurrentRow.Item("MilesToCoop")), 0.0, CurrentRow.Item("MilesToCoop"))
      MilesToDreams__Numeric = IIf(IsDBNull(CurrentRow.Item("MilesToDreams")), 0.0, CurrentRow.Item("MilesToDreams"))
      Phone__String = IIf(IsDBNull(CurrentRow.Item("Phone")), "", CurrentRow.Item("Phone"))
      Sleeps__Integer = IIf(IsDBNull(CurrentRow.Item("Sleeps")), 0, CurrentRow.Item("Sleeps"))
      Bedrooms__Numeric = IIf(IsDBNull(CurrentRow.Item("Bedrooms")), 0.0, CurrentRow.Item("Bedrooms"))
      Baths__Numeric = IIf(IsDBNull(CurrentRow.Item("Baths")), 0.0, CurrentRow.Item("Baths"))
      Host_ID__Integer = IIf(IsDBNull(CurrentRow.Item("Host_ID")), 0, CurrentRow.Item("Host_ID"))
      SummerRate__Numeric = IIf(IsDBNull(CurrentRow.Item("SummerRate")), 0.0, CurrentRow.Item("SummerRate"))
      DamageDeposit__Numeric = IIf(IsDBNull(CurrentRow.Item("DamageDeposit")), 0.0, CurrentRow.Item("DamageDeposit"))
      CableTV__Integer = IIf(IsDBNull(CurrentRow.Item("CableTV")), 0, CurrentRow.Item("CableTV"))
      VCR__Integer = IIf(IsDBNull(CurrentRow.Item("VCR")), 0, CurrentRow.Item("VCR"))
      Grill__Integer = IIf(IsDBNull(CurrentRow.Item("Grill")), 0, CurrentRow.Item("Grill"))
      WindowAC__Integer = IIf(IsDBNull(CurrentRow.Item("WindowAC")), 0, CurrentRow.Item("WindowAC"))
      WasherDryer__Integer = IIf(IsDBNull(CurrentRow.Item("WasherDryer")), 0, CurrentRow.Item("WasherDryer"))
      Telephone__Integer = IIf(IsDBNull(CurrentRow.Item("Telephone")), 0, CurrentRow.Item("Telephone"))
      Broadband__Integer = IIf(IsDBNull(CurrentRow.Item("Broadband")), 0, CurrentRow.Item("Broadband"))
      Handicap__Integer = IIf(IsDBNull(CurrentRow.Item("Handicap")), 0, CurrentRow.Item("Handicap"))
      WheelchairAccess__Integer = IIf(IsDBNull(CurrentRow.Item("WheelchairAccess")), 0, CurrentRow.Item("WheelchairAccess"))
      PropertyNotes__String = IIf(IsDBNull(CurrentRow.Item("PropertyNotes")), "", CurrentRow.Item("PropertyNotes"))
      Commission_Pct__Integer = IIf(IsDBNull(CurrentRow.Item("Commission_Pct")), 0, CurrentRow.Item("Commission_Pct"))
      TaxRate__Integer = IIf(IsDBNull(CurrentRow.Item("TaxRate")), 0, CurrentRow.Item("TaxRate"))
      CollectTax__String = IIf(IsDBNull(CurrentRow.Item("CollectTax")), "", CurrentRow.Item("CollectTax"))
      TaxIDNumber__String = IIf(IsDBNull(CurrentRow.Item("TaxIDNumber")), "", CurrentRow.Item("TaxIDNumber"))
      Status__String = IIf(IsDBNull(CurrentRow.Item("Status")), "", CurrentRow.Item("Status"))
      WebPage__String = IIf(IsDBNull(CurrentRow.Item("WebPage")), "", CurrentRow.Item("WebPage"))
      Longitude__Numeric = IIf(IsDBNull(CurrentRow.Item("Longitude")), 0.0, CurrentRow.Item("Longitude"))
      Latitude__Numeric = IIf(IsDBNull(CurrentRow.Item("Latitude")), 0.0, CurrentRow.Item("Latitude"))
      Distance2CDP__Numeric = IIf(IsDBNull(CurrentRow.Item("Distance2CDP")), 0.0, CurrentRow.Item("Distance2CDP"))
      Distance2BW__Numeric = IIf(IsDBNull(CurrentRow.Item("Distance2BW")), 0.0, CurrentRow.Item("Distance2BW"))
      Distance2ASV__Numeric = IIf(IsDBNull(CurrentRow.Item("Distance2ASV")), 0.0, CurrentRow.Item("Distance2ASV"))
      Distance2Coop__Numeric = IIf(IsDBNull(CurrentRow.Item("Distance2Coop")), 0.0, CurrentRow.Item("Distance2Coop"))
      DVD__Integer = IIf(IsDBNull(CurrentRow.Item("DVD")), 0, CurrentRow.Item("DVD"))
      CentralAC__Integer = IIf(IsDBNull(CurrentRow.Item("CentralAC")), 0, CurrentRow.Item("CentralAC"))
      DishSatellite__Integer = IIf(IsDBNull(CurrentRow.Item("DishSatellite")), 0, CurrentRow.Item("DishSatellite"))
      Dialup__Integer = IIf(IsDBNull(CurrentRow.Item("Dialup")), 0, CurrentRow.Item("Dialup"))
      SquareFootage__Integer = IIf(IsDBNull(CurrentRow.Item("SquareFootage")), 0, CurrentRow.Item("SquareFootage"))
      KingBeds__Integer = IIf(IsDBNull(CurrentRow.Item("KingBeds")), 0, CurrentRow.Item("KingBeds"))
      QueenBeds__Integer = IIf(IsDBNull(CurrentRow.Item("QueenBeds")), 0, CurrentRow.Item("QueenBeds"))
      DoubleBeds__Integer = IIf(IsDBNull(CurrentRow.Item("DoubleBeds")), 0, CurrentRow.Item("DoubleBeds"))
      TwinBeds__Integer = IIf(IsDBNull(CurrentRow.Item("TwinBeds")), 0, CurrentRow.Item("TwinBeds"))
      BunkBeds__Integer = IIf(IsDBNull(CurrentRow.Item("BunkBeds")), 0, CurrentRow.Item("BunkBeds"))
      SleeperSofa__Integer = IIf(IsDBNull(CurrentRow.Item("SleeperSofa")), 0, CurrentRow.Item("SleeperSofa"))
      Futon__Integer = IIf(IsDBNull(CurrentRow.Item("Futon")), 0, CurrentRow.Item("Futon"))
      Dishwasher__Integer = IIf(IsDBNull(CurrentRow.Item("Dishwasher")), 0, CurrentRow.Item("Dishwasher"))
      Pool__Integer = IIf(IsDBNull(CurrentRow.Item("Pool")), 0, CurrentRow.Item("Pool"))
      Wireless__Integer = IIf(IsDBNull(CurrentRow.Item("Wireless")), 0, CurrentRow.Item("Wireless"))
      PrivatePond__Integer = IIf(IsDBNull(CurrentRow.Item("PrivatePond")), 0, CurrentRow.Item("PrivatePond"))
      TeamParties__Integer = IIf(IsDBNull(CurrentRow.Item("TeamParties")), 0, CurrentRow.Item("TeamParties"))
      WebImage__String = IIf(IsDBNull(CurrentRow.Item("WebImage")), "", CurrentRow.Item("WebImage"))
      PropertyGroup__Integer = IIf(IsDBNull(CurrentRow.Item("PropertyGroup")), 0, CurrentRow.Item("PropertyGroup"))
      QBListID__String = IIf(IsDBNull(CurrentRow.Item("QBListID")), "", CurrentRow.Item("QBListID"))
      CheckName__String = IIf(IsDBNull(CurrentRow.Item("CheckName")), "", CurrentRow.Item("CheckName"))
      DiscountedRate__Numeric = IIf(IsDBNull(CurrentRow.Item("DiscountedRate")), 0.0, CurrentRow.Item("DiscountedRate"))
      WebMasterDescription__String = IIf(IsDBNull(CurrentRow.Item("WebMasterDescription")), "", CurrentRow.Item("WebMasterDescription"))
      WebDetailsDescription__String = IIf(IsDBNull(CurrentRow.Item("WebDetailsDescription")), "", CurrentRow.Item("WebDetailsDescription"))
      WebMasterTitle__String = IIf(IsDBNull(CurrentRow.Item("WebMasterTitle")), "", CurrentRow.Item("WebMasterTitle"))
      WebDetailsTitle__String = IIf(IsDBNull(CurrentRow.Item("WebDetailsTitle")), "", CurrentRow.Item("WebDetailsTitle"))
      WebDetailsLeftSection__String = IIf(IsDBNull(CurrentRow.Item("WebDetailsLeftSection")), "", CurrentRow.Item("WebDetailsLeftSection"))
      WebDetailsRightSection__String = IIf(IsDBNull(CurrentRow.Item("WebDetailsRightSection")), "", CurrentRow.Item("WebDetailsRightSection"))
      Category_ID2__Integer = IIf(IsDBNull(CurrentRow.Item("Category_ID2")), 0, CurrentRow.Item("Category_ID2"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [Properties](")
    If sInsUpdCategory_ID.ToString <> "" Then
      sb.Append("Category_ID,")
      sbv.Append(sInsUpdCategory_ID & ",")
    End If
    If sInsUpdPropertyName.ToString <> "" Then
      sb.Append("PropertyName,")
      sbv.Append(sInsUpdPropertyName & ",")
    End If
    If sInsUpdAddress.ToString <> "" Then
      sb.Append("Address,")
      sbv.Append(sInsUpdAddress & ",")
    End If
    If sInsUpdAddress2.ToString <> "" Then
      sb.Append("Address2,")
      sbv.Append(sInsUpdAddress2 & ",")
    End If
    If sInsUpdCity.ToString <> "" Then
      sb.Append("City,")
      sbv.Append(sInsUpdCity & ",")
    End If
    If sInsUpdZip.ToString <> "" Then
      sb.Append("Zip,")
      sbv.Append(sInsUpdZip & ",")
    End If
    If sInsUpdMilesToCoop.ToString <> "" Then
      sb.Append("MilesToCoop,")
      sbv.Append(sInsUpdMilesToCoop & ",")
    End If
    If sInsUpdMilesToDreams.ToString <> "" Then
      sb.Append("MilesToDreams,")
      sbv.Append(sInsUpdMilesToDreams & ",")
    End If
    If sInsUpdPhone.ToString <> "" Then
      sb.Append("Phone,")
      sbv.Append(sInsUpdPhone & ",")
    End If
    If sInsUpdSleeps.ToString <> "" Then
      sb.Append("Sleeps,")
      sbv.Append(sInsUpdSleeps & ",")
    End If
    If sInsUpdBedrooms.ToString <> "" Then
      sb.Append("Bedrooms,")
      sbv.Append(sInsUpdBedrooms & ",")
    End If
    If sInsUpdBaths.ToString <> "" Then
      sb.Append("Baths,")
      sbv.Append(sInsUpdBaths & ",")
    End If
    If sInsUpdHost_ID.ToString <> "" Then
      sb.Append("Host_ID,")
      sbv.Append(sInsUpdHost_ID & ",")
    End If
    If sInsUpdSummerRate.ToString <> "" Then
      sb.Append("SummerRate,")
      sbv.Append(sInsUpdSummerRate & ",")
    End If
    If sInsUpdDamageDeposit.ToString <> "" Then
      sb.Append("DamageDeposit,")
      sbv.Append(sInsUpdDamageDeposit & ",")
    End If
    If sInsUpdCableTV.ToString <> "" Then
      sb.Append("CableTV,")
      sbv.Append(sInsUpdCableTV & ",")
    End If
    If sInsUpdVCR.ToString <> "" Then
      sb.Append("VCR,")
      sbv.Append(sInsUpdVCR & ",")
    End If
    If sInsUpdGrill.ToString <> "" Then
      sb.Append("Grill,")
      sbv.Append(sInsUpdGrill & ",")
    End If
    If sInsUpdWindowAC.ToString <> "" Then
      sb.Append("WindowAC,")
      sbv.Append(sInsUpdWindowAC & ",")
    End If
    If sInsUpdWasherDryer.ToString <> "" Then
      sb.Append("WasherDryer,")
      sbv.Append(sInsUpdWasherDryer & ",")
    End If
    If sInsUpdTelephone.ToString <> "" Then
      sb.Append("Telephone,")
      sbv.Append(sInsUpdTelephone & ",")
    End If
    If sInsUpdBroadband.ToString <> "" Then
      sb.Append("Broadband,")
      sbv.Append(sInsUpdBroadband & ",")
    End If
    If sInsUpdHandicap.ToString <> "" Then
      sb.Append("Handicap,")
      sbv.Append(sInsUpdHandicap & ",")
    End If
    If sInsUpdWheelchairAccess.ToString <> "" Then
      sb.Append("WheelchairAccess,")
      sbv.Append(sInsUpdWheelchairAccess & ",")
    End If
    If sInsUpdPropertyNotes.ToString <> "" Then
      sb.Append("PropertyNotes,")
      sbv.Append(sInsUpdPropertyNotes & ",")
    End If
    If sInsUpdCommission_Pct.ToString <> "" Then
      sb.Append("Commission_Pct,")
      sbv.Append(sInsUpdCommission_Pct & ",")
    End If
    If sInsUpdTaxRate.ToString <> "" Then
      sb.Append("TaxRate,")
      sbv.Append(sInsUpdTaxRate & ",")
    End If
    If sInsUpdCollectTax.ToString <> "" Then
      sb.Append("CollectTax,")
      sbv.Append(sInsUpdCollectTax & ",")
    End If
    If sInsUpdTaxIDNumber.ToString <> "" Then
      sb.Append("TaxIDNumber,")
      sbv.Append(sInsUpdTaxIDNumber & ",")
    End If
    If sInsUpdStatus.ToString <> "" Then
      sb.Append("Status,")
      sbv.Append(sInsUpdStatus & ",")
    End If
    If sInsUpdWebPage.ToString <> "" Then
      sb.Append("WebPage,")
      sbv.Append(sInsUpdWebPage & ",")
    End If
    If sInsUpdLongitude.ToString <> "" Then
      sb.Append("Longitude,")
      sbv.Append(sInsUpdLongitude & ",")
    End If
    If sInsUpdLatitude.ToString <> "" Then
      sb.Append("Latitude,")
      sbv.Append(sInsUpdLatitude & ",")
    End If
    If sInsUpdDistance2CDP.ToString <> "" Then
      sb.Append("Distance2CDP,")
      sbv.Append(sInsUpdDistance2CDP & ",")
    End If
    If sInsUpdDistance2BW.ToString <> "" Then
      sb.Append("Distance2BW,")
      sbv.Append(sInsUpdDistance2BW & ",")
    End If
    If sInsUpdDistance2ASV.ToString <> "" Then
      sb.Append("Distance2ASV,")
      sbv.Append(sInsUpdDistance2ASV & ",")
    End If
    If sInsUpdDistance2Coop.ToString <> "" Then
      sb.Append("Distance2Coop,")
      sbv.Append(sInsUpdDistance2Coop & ",")
    End If
    If sInsUpdDVD.ToString <> "" Then
      sb.Append("DVD,")
      sbv.Append(sInsUpdDVD & ",")
    End If
    If sInsUpdCentralAC.ToString <> "" Then
      sb.Append("CentralAC,")
      sbv.Append(sInsUpdCentralAC & ",")
    End If
    If sInsUpdDishSatellite.ToString <> "" Then
      sb.Append("DishSatellite,")
      sbv.Append(sInsUpdDishSatellite & ",")
    End If
    If sInsUpdDialup.ToString <> "" Then
      sb.Append("Dialup,")
      sbv.Append(sInsUpdDialup & ",")
    End If
    If sInsUpdSquareFootage.ToString <> "" Then
      sb.Append("SquareFootage,")
      sbv.Append(sInsUpdSquareFootage & ",")
    End If
    If sInsUpdKingBeds.ToString <> "" Then
      sb.Append("KingBeds,")
      sbv.Append(sInsUpdKingBeds & ",")
    End If
    If sInsUpdQueenBeds.ToString <> "" Then
      sb.Append("QueenBeds,")
      sbv.Append(sInsUpdQueenBeds & ",")
    End If
    If sInsUpdDoubleBeds.ToString <> "" Then
      sb.Append("DoubleBeds,")
      sbv.Append(sInsUpdDoubleBeds & ",")
    End If
    If sInsUpdTwinBeds.ToString <> "" Then
      sb.Append("TwinBeds,")
      sbv.Append(sInsUpdTwinBeds & ",")
    End If
    If sInsUpdBunkBeds.ToString <> "" Then
      sb.Append("BunkBeds,")
      sbv.Append(sInsUpdBunkBeds & ",")
    End If
    If sInsUpdSleeperSofa.ToString <> "" Then
      sb.Append("SleeperSofa,")
      sbv.Append(sInsUpdSleeperSofa & ",")
    End If
    If sInsUpdFuton.ToString <> "" Then
      sb.Append("Futon,")
      sbv.Append(sInsUpdFuton & ",")
    End If
    If sInsUpdDishwasher.ToString <> "" Then
      sb.Append("Dishwasher,")
      sbv.Append(sInsUpdDishwasher & ",")
    End If
    If sInsUpdPool.ToString <> "" Then
      sb.Append("Pool,")
      sbv.Append(sInsUpdPool & ",")
    End If
    If sInsUpdWireless.ToString <> "" Then
      sb.Append("Wireless,")
      sbv.Append(sInsUpdWireless & ",")
    End If
    If sInsUpdPrivatePond.ToString <> "" Then
      sb.Append("PrivatePond,")
      sbv.Append(sInsUpdPrivatePond & ",")
    End If
    If sInsUpdTeamParties.ToString <> "" Then
      sb.Append("TeamParties,")
      sbv.Append(sInsUpdTeamParties & ",")
    End If
    If sInsUpdWebImage.ToString <> "" Then
      sb.Append("WebImage,")
      sbv.Append(sInsUpdWebImage & ",")
    End If
    If sInsUpdPropertyGroup.ToString <> "" Then
      sb.Append("PropertyGroup,")
      sbv.Append(sInsUpdPropertyGroup & ",")
    End If
    If sInsUpdQBListID.ToString <> "" Then
      sb.Append("QBListID,")
      sbv.Append(sInsUpdQBListID & ",")
    End If
    If sInsUpdCheckName.ToString <> "" Then
      sb.Append("CheckName,")
      sbv.Append(sInsUpdCheckName & ",")
    End If
    If sInsUpdDiscountedRate.ToString <> "" Then
      sb.Append("DiscountedRate,")
      sbv.Append(sInsUpdDiscountedRate & ",")
    End If
    If sInsUpdWebMasterDescription.ToString <> "" Then
      sb.Append("WebMasterDescription,")
      sbv.Append(sInsUpdWebMasterDescription & ",")
    End If
    If sInsUpdWebDetailsDescription.ToString <> "" Then
      sb.Append("WebDetailsDescription,")
      sbv.Append(sInsUpdWebDetailsDescription & ",")
    End If
    If sInsUpdWebMasterTitle.ToString <> "" Then
      sb.Append("WebMasterTitle,")
      sbv.Append(sInsUpdWebMasterTitle & ",")
    End If
    If sInsUpdWebDetailsTitle.ToString <> "" Then
      sb.Append("WebDetailsTitle,")
      sbv.Append(sInsUpdWebDetailsTitle & ",")
    End If
    If sInsUpdWebDetailsLeftSection.ToString <> "" Then
      sb.Append("WebDetailsLeftSection,")
      sbv.Append(sInsUpdWebDetailsLeftSection & ",")
    End If
    If sInsUpdWebDetailsRightSection.ToString <> "" Then
      sb.Append("WebDetailsRightSection,")
      sbv.Append(sInsUpdWebDetailsRightSection & ",")
    End If
    If sInsUpdCategory_ID2.ToString <> "" Then
      sb.Append("Category_ID2,")
      sbv.Append(sInsUpdCategory_ID2 & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(Property_ID) from [Properties]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    Property_ID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [Properties] Set ")
    If sInsUpdCategory_ID.ToString <> "" Then sb.Append("Category_ID=" & sInsUpdCategory_ID & ",")
    If sInsUpdPropertyName.ToString <> "" Then sb.Append("PropertyName=" & sInsUpdPropertyName & ",")
    If sInsUpdAddress.ToString <> "" Then sb.Append("Address=" & sInsUpdAddress & ",")
    If sInsUpdAddress2.ToString <> "" Then sb.Append("Address2=" & sInsUpdAddress2 & ",")
    If sInsUpdCity.ToString <> "" Then sb.Append("City=" & sInsUpdCity & ",")
    If sInsUpdZip.ToString <> "" Then sb.Append("Zip=" & sInsUpdZip & ",")
    If sInsUpdMilesToCoop.ToString <> "" Then sb.Append("MilesToCoop=" & sInsUpdMilesToCoop & ",")
    If sInsUpdMilesToDreams.ToString <> "" Then sb.Append("MilesToDreams=" & sInsUpdMilesToDreams & ",")
    If sInsUpdPhone.ToString <> "" Then sb.Append("Phone=" & sInsUpdPhone & ",")
    If sInsUpdSleeps.ToString <> "" Then sb.Append("Sleeps=" & sInsUpdSleeps & ",")
    If sInsUpdBedrooms.ToString <> "" Then sb.Append("Bedrooms=" & sInsUpdBedrooms & ",")
    If sInsUpdBaths.ToString <> "" Then sb.Append("Baths=" & sInsUpdBaths & ",")
    If sInsUpdHost_ID.ToString <> "" Then sb.Append("Host_ID=" & sInsUpdHost_ID & ",")
    If sInsUpdSummerRate.ToString <> "" Then sb.Append("SummerRate=" & sInsUpdSummerRate & ",")
    If sInsUpdDamageDeposit.ToString <> "" Then sb.Append("DamageDeposit=" & sInsUpdDamageDeposit & ",")
    If sInsUpdCableTV.ToString <> "" Then sb.Append("CableTV=" & sInsUpdCableTV & ",")
    If sInsUpdVCR.ToString <> "" Then sb.Append("VCR=" & sInsUpdVCR & ",")
    If sInsUpdGrill.ToString <> "" Then sb.Append("Grill=" & sInsUpdGrill & ",")
    If sInsUpdWindowAC.ToString <> "" Then sb.Append("WindowAC=" & sInsUpdWindowAC & ",")
    If sInsUpdWasherDryer.ToString <> "" Then sb.Append("WasherDryer=" & sInsUpdWasherDryer & ",")
    If sInsUpdTelephone.ToString <> "" Then sb.Append("Telephone=" & sInsUpdTelephone & ",")
    If sInsUpdBroadband.ToString <> "" Then sb.Append("Broadband=" & sInsUpdBroadband & ",")
    If sInsUpdHandicap.ToString <> "" Then sb.Append("Handicap=" & sInsUpdHandicap & ",")
    If sInsUpdWheelchairAccess.ToString <> "" Then sb.Append("WheelchairAccess=" & sInsUpdWheelchairAccess & ",")
    If sInsUpdPropertyNotes.ToString <> "" Then sb.Append("PropertyNotes=" & sInsUpdPropertyNotes & ",")
    If sInsUpdCommission_Pct.ToString <> "" Then sb.Append("Commission_Pct=" & sInsUpdCommission_Pct & ",")
    If sInsUpdTaxRate.ToString <> "" Then sb.Append("TaxRate=" & sInsUpdTaxRate & ",")
    If sInsUpdCollectTax.ToString <> "" Then sb.Append("CollectTax=" & sInsUpdCollectTax & ",")
    If sInsUpdTaxIDNumber.ToString <> "" Then sb.Append("TaxIDNumber=" & sInsUpdTaxIDNumber & ",")
    If sInsUpdStatus.ToString <> "" Then sb.Append("Status=" & sInsUpdStatus & ",")
    If sInsUpdWebPage.ToString <> "" Then sb.Append("WebPage=" & sInsUpdWebPage & ",")
    If sInsUpdLongitude.ToString <> "" Then sb.Append("Longitude=" & sInsUpdLongitude & ",")
    If sInsUpdLatitude.ToString <> "" Then sb.Append("Latitude=" & sInsUpdLatitude & ",")
    If sInsUpdDistance2CDP.ToString <> "" Then sb.Append("Distance2CDP=" & sInsUpdDistance2CDP & ",")
    If sInsUpdDistance2BW.ToString <> "" Then sb.Append("Distance2BW=" & sInsUpdDistance2BW & ",")
    If sInsUpdDistance2ASV.ToString <> "" Then sb.Append("Distance2ASV=" & sInsUpdDistance2ASV & ",")
    If sInsUpdDistance2Coop.ToString <> "" Then sb.Append("Distance2Coop=" & sInsUpdDistance2Coop & ",")
    If sInsUpdDVD.ToString <> "" Then sb.Append("DVD=" & sInsUpdDVD & ",")
    If sInsUpdCentralAC.ToString <> "" Then sb.Append("CentralAC=" & sInsUpdCentralAC & ",")
    If sInsUpdDishSatellite.ToString <> "" Then sb.Append("DishSatellite=" & sInsUpdDishSatellite & ",")
    If sInsUpdDialup.ToString <> "" Then sb.Append("Dialup=" & sInsUpdDialup & ",")
    If sInsUpdSquareFootage.ToString <> "" Then sb.Append("SquareFootage=" & sInsUpdSquareFootage & ",")
    If sInsUpdKingBeds.ToString <> "" Then sb.Append("KingBeds=" & sInsUpdKingBeds & ",")
    If sInsUpdQueenBeds.ToString <> "" Then sb.Append("QueenBeds=" & sInsUpdQueenBeds & ",")
    If sInsUpdDoubleBeds.ToString <> "" Then sb.Append("DoubleBeds=" & sInsUpdDoubleBeds & ",")
    If sInsUpdTwinBeds.ToString <> "" Then sb.Append("TwinBeds=" & sInsUpdTwinBeds & ",")
    If sInsUpdBunkBeds.ToString <> "" Then sb.Append("BunkBeds=" & sInsUpdBunkBeds & ",")
    If sInsUpdSleeperSofa.ToString <> "" Then sb.Append("SleeperSofa=" & sInsUpdSleeperSofa & ",")
    If sInsUpdFuton.ToString <> "" Then sb.Append("Futon=" & sInsUpdFuton & ",")
    If sInsUpdDishwasher.ToString <> "" Then sb.Append("Dishwasher=" & sInsUpdDishwasher & ",")
    If sInsUpdPool.ToString <> "" Then sb.Append("Pool=" & sInsUpdPool & ",")
    If sInsUpdWireless.ToString <> "" Then sb.Append("Wireless=" & sInsUpdWireless & ",")
    If sInsUpdPrivatePond.ToString <> "" Then sb.Append("PrivatePond=" & sInsUpdPrivatePond & ",")
    If sInsUpdTeamParties.ToString <> "" Then sb.Append("TeamParties=" & sInsUpdTeamParties & ",")
    If sInsUpdWebImage.ToString <> "" Then sb.Append("WebImage=" & sInsUpdWebImage & ",")
    If sInsUpdPropertyGroup.ToString <> "" Then sb.Append("PropertyGroup=" & sInsUpdPropertyGroup & ",")
    If sInsUpdQBListID.ToString <> "" Then sb.Append("QBListID=" & sInsUpdQBListID & ",")
    If sInsUpdCheckName.ToString <> "" Then sb.Append("CheckName=" & sInsUpdCheckName & ",")
    If sInsUpdDiscountedRate.ToString <> "" Then sb.Append("DiscountedRate=" & sInsUpdDiscountedRate & ",")
    If sInsUpdWebMasterDescription.ToString <> "" Then sb.Append("WebMasterDescription=" & sInsUpdWebMasterDescription & ",")
    If sInsUpdWebDetailsDescription.ToString <> "" Then sb.Append("WebDetailsDescription=" & sInsUpdWebDetailsDescription & ",")
    If sInsUpdWebMasterTitle.ToString <> "" Then sb.Append("WebMasterTitle=" & sInsUpdWebMasterTitle & ",")
    If sInsUpdWebDetailsTitle.ToString <> "" Then sb.Append("WebDetailsTitle=" & sInsUpdWebDetailsTitle & ",")
    If sInsUpdWebDetailsLeftSection.ToString <> "" Then sb.Append("WebDetailsLeftSection=" & sInsUpdWebDetailsLeftSection & ",")
    If sInsUpdWebDetailsRightSection.ToString <> "" Then sb.Append("WebDetailsRightSection=" & sInsUpdWebDetailsRightSection & ",")
    If sInsUpdCategory_ID2.ToString <> "" Then sb.Append("Category_ID2=" & sInsUpdCategory_ID2 & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where Property_ID=" & sInsUpdProperty_ID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [Properties] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("Property_ID=" & sInsUpdProperty_ID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TableBookingsContactLog

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iLogID As Int32
  Private sInsUpdLogID As String
  Property LogID_PK__Integer() As Int32
    Get
      Return iLogID
    End Get
    Set(ByVal Value As Int32)
      iLogID = Value
      sInsUpdLogID = oUtil.FixParam(iLogID, True)
    End Set
  End Property

  Private iBookingID As Int32
  Private sInsUpdBookingID As String
  Property BookingID__Integer() As Int32
    Get
      Return iBookingID
    End Get
    Set(ByVal Value As Int32)
      iBookingID = Value
      sInsUpdBookingID = oUtil.FixParam(iBookingID, True)
    End Set
  End Property

  Private sLogDate As String
  Private sInsUpdLogDate As String
  Property LogDate__Date() As String
    Get
      Return sLogDate
    End Get
    Set(ByVal Value As String)
      sLogDate = Value
      sInsUpdLogDate = oUtil.FixParam(sLogDate, True)
    End Set
  End Property

  Private sUser As String
  Private sInsUpdUser As String
  Property User__String() As String
    Get
      Return sUser
    End Get
    Set(ByVal Value As String)
      sUser = Value
      sInsUpdUser = oUtil.FixParam(sUser, True)
    End Set
  End Property

  Private sLogEntry As String
  Private sInsUpdLogEntry As String
  Property LogEntry__String() As String
    Get
      Return sLogEntry
    End Get
    Set(ByVal Value As String)
      sLogEntry = Value
      sInsUpdLogEntry = oUtil.FixParam(sLogEntry, True)
    End Set
  End Property

  Private sLogUpdates As String
  Private sInsUpdLogUpdates As String
  Property LogUpdates__String() As String
    Get
      Return sLogUpdates
    End Get
    Set(ByVal Value As String)
      sLogUpdates = Value
      sInsUpdLogUpdates = oUtil.FixParam(sLogUpdates, True)
    End Set
  End Property

  Private sEmailAddress As String
  Private sInsUpdEmailAddress As String
  Property EmailAddress__String() As String
    Get
      Return sEmailAddress
    End Get
    Set(ByVal Value As String)
      sEmailAddress = Value
      sInsUpdEmailAddress = oUtil.FixParam(sEmailAddress, True)
    End Set
  End Property

  Private sEmailType As String
  Private sInsUpdEmailType As String
  Property EmailType__String() As String
    Get
      Return sEmailType
    End Get
    Set(ByVal Value As String)
      sEmailType = Value
      sInsUpdEmailType = oUtil.FixParam(sEmailType, True)
    End Set
  End Property

  Private sLetterType As String
  Private sInsUpdLetterType As String
  Property LetterType__String() As String
    Get
      Return sLetterType
    End Get
    Set(ByVal Value As String)
      sLetterType = Value
      sInsUpdLetterType = oUtil.FixParam(sLetterType, True)
    End Set
  End Property

  Private sEmailSent As String
  Private sInsUpdEmailSent As String
  Property EmailSent__String() As String
    Get
      Return sEmailSent
    End Get
    Set(ByVal Value As String)
      sEmailSent = Value
      sInsUpdEmailSent = oUtil.FixParam(sEmailSent, True)
    End Set
  End Property

  Private sEmailSubject As String
  Private sInsUpdEmailSubject As String
  Property EmailSubject__String() As String
    Get
      Return sEmailSubject
    End Get
    Set(ByVal Value As String)
      sEmailSubject = Value
      sInsUpdEmailSubject = oUtil.FixParam(sEmailSubject, True)
    End Set
  End Property

  Private sEmailContentType As String
  Private sInsUpdEmailContentType As String
  Property EmailContentType__String() As String
    Get
      Return sEmailContentType
    End Get
    Set(ByVal Value As String)
      sEmailContentType = Value
      sInsUpdEmailContentType = oUtil.FixParam(sEmailContentType, True)
    End Set
  End Property

  Public Sub Clear()
    iLogID = 0
    sInsUpdLogID = ""
    iBookingID = 0
    sInsUpdBookingID = ""
    sLogDate = ""
    sInsUpdLogDate = ""
    sUser = ""
    sInsUpdUser = ""
    sLogEntry = ""
    sInsUpdLogEntry = ""
    sLogUpdates = ""
    sInsUpdLogUpdates = ""
    sEmailAddress = ""
    sInsUpdEmailAddress = ""
    sEmailType = ""
    sInsUpdEmailType = ""
    sLetterType = ""
    sInsUpdLetterType = ""
    sEmailSent = ""
    sInsUpdEmailSent = ""
    sEmailSubject = ""
    sInsUpdEmailSubject = ""
    sEmailContentType = ""
    sInsUpdEmailContentType = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdLogID.ToString = "'-12345'" Then sb.Append("LogID,")
        If sInsUpdBookingID.ToString = "'-12345'" Then sb.Append("BookingID,")
        If sInsUpdLogDate.ToString = "'Select'" Then sb.Append("LogDate,")
        If sInsUpdUser.ToString = "'Select'" Then sb.Append("[User],")
        If sInsUpdLogEntry.ToString = "'Select'" Then sb.Append("LogEntry,")
        If sInsUpdLogUpdates.ToString = "'Select'" Then sb.Append("LogUpdates,")
        If sInsUpdEmailAddress.ToString = "'Select'" Then sb.Append("EmailAddress,")
        If sInsUpdEmailType.ToString = "'Select'" Then sb.Append("EmailType,")
        If sInsUpdLetterType.ToString = "'Select'" Then sb.Append("LetterType,")
        If sInsUpdEmailSent.ToString = "'Select'" Then sb.Append("EmailSent,")
        If sInsUpdEmailSubject.ToString = "'Select'" Then sb.Append("EmailSubject,")
        If sInsUpdEmailContentType.ToString = "'Select'" Then sb.Append("EmailContentType,")
      Else
        sb.Append("LogID,")
        sb.Append("BookingID,")
        sb.Append("LogDate,")
        sb.Append("[User],")
        sb.Append("LogEntry,")
        sb.Append("LogUpdates,")
        sb.Append("EmailAddress,")
        sb.Append("EmailType,")
        sb.Append("LetterType,")
        sb.Append("EmailSent,")
        sb.Append("EmailSubject,")
        sb.Append("EmailContentType,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [BookingsContactLog]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdLogID.ToString <> "") And (sInsUpdLogID <> "'-12345'") Then sbw.Append("LogID=" & sInsUpdLogID & " and ")
      If (sInsUpdBookingID.ToString <> "") And (sInsUpdBookingID <> "'-12345'") Then sbw.Append("BookingID=" & sInsUpdBookingID & " and ")
      If (sInsUpdLogDate.ToString <> "") And (sInsUpdLogDate <> "'Select'") Then sbw.Append("LogDate=" & sInsUpdLogDate & " and ")
      If (sInsUpdUser.ToString <> "") And (sInsUpdUser <> "'Select'") Then sbw.Append("[User]=" & sInsUpdUser & " and ")
      If (sInsUpdLogEntry.ToString <> "") And (sInsUpdLogEntry <> "'Select'") Then sbw.Append("LogEntry=" & sInsUpdLogEntry & " and ")
      If (sInsUpdLogUpdates.ToString <> "") And (sInsUpdLogUpdates <> "'Select'") Then sbw.Append("LogUpdates=" & sInsUpdLogUpdates & " and ")
      If (sInsUpdEmailAddress.ToString <> "") And (sInsUpdEmailAddress <> "'Select'") Then sbw.Append("EmailAddress=" & sInsUpdEmailAddress & " and ")
      If (sInsUpdEmailType.ToString <> "") And (sInsUpdEmailType <> "'Select'") Then sbw.Append("EmailType=" & sInsUpdEmailType & " and ")
      If (sInsUpdLetterType.ToString <> "") And (sInsUpdLetterType <> "'Select'") Then sbw.Append("LetterType=" & sInsUpdLetterType & " and ")
      If (sInsUpdEmailSent.ToString <> "") And (sInsUpdEmailSent <> "'Select'") Then sbw.Append("EmailSent=" & sInsUpdEmailSent & " and ")
      If (sInsUpdEmailSubject.ToString <> "") And (sInsUpdEmailSubject <> "'Select'") Then sbw.Append("EmailSubject=" & sInsUpdEmailSubject & " and ")
      If (sInsUpdEmailContentType.ToString <> "") And (sInsUpdEmailContentType <> "'Select'") Then sbw.Append("EmailContentType=" & sInsUpdEmailContentType & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      LogID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("LogID")), 0, CurrentRow.Item("LogID").ToString)
      BookingID__Integer = IIf(IsDBNull(CurrentRow.Item("BookingID")), 0, CurrentRow.Item("BookingID"))
      LogDate__Date = IIf(IsDBNull(CurrentRow.Item("LogDate")), "", CurrentRow.Item("LogDate"))
      User__String = IIf(IsDBNull(CurrentRow.Item("User")), "", CurrentRow.Item("User"))
      LogEntry__String = IIf(IsDBNull(CurrentRow.Item("LogEntry")), "", CurrentRow.Item("LogEntry"))
      LogUpdates__String = IIf(IsDBNull(CurrentRow.Item("LogUpdates")), "", CurrentRow.Item("LogUpdates"))
      EmailAddress__String = IIf(IsDBNull(CurrentRow.Item("EmailAddress")), "", CurrentRow.Item("EmailAddress"))
      EmailType__String = IIf(IsDBNull(CurrentRow.Item("EmailType")), "", CurrentRow.Item("EmailType"))
      LetterType__String = IIf(IsDBNull(CurrentRow.Item("LetterType")), "", CurrentRow.Item("LetterType"))
      EmailSent__String = IIf(IsDBNull(CurrentRow.Item("EmailSent")), "", CurrentRow.Item("EmailSent"))
      EmailSubject__String = IIf(IsDBNull(CurrentRow.Item("EmailSubject")), "", CurrentRow.Item("EmailSubject"))
      EmailContentType__String = IIf(IsDBNull(CurrentRow.Item("EmailContentType")), "", CurrentRow.Item("EmailContentType"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [BookingsContactLog](")
    If sInsUpdBookingID.ToString <> "" Then
      sb.Append("BookingID,")
      sbv.Append(sInsUpdBookingID & ",")
    End If
    If sInsUpdLogDate.ToString <> "" Then
      sb.Append("LogDate,")
      sbv.Append(sInsUpdLogDate & ",")
    End If
    If sInsUpdUser.ToString <> "" Then
      sb.Append("[User],")
      sbv.Append(sInsUpdUser & ",")
    End If
    If sInsUpdLogEntry.ToString <> "" Then
      sb.Append("LogEntry,")
      sbv.Append(sInsUpdLogEntry & ",")
    End If
    If sInsUpdLogUpdates.ToString <> "" Then
      sb.Append("LogUpdates,")
      sbv.Append(sInsUpdLogUpdates & ",")
    End If
    If sInsUpdEmailAddress.ToString <> "" Then
      sb.Append("EmailAddress,")
      sbv.Append(sInsUpdEmailAddress & ",")
    End If
    If sInsUpdEmailType.ToString <> "" Then
      sb.Append("EmailType,")
      sbv.Append(sInsUpdEmailType & ",")
    End If
    If sInsUpdLetterType.ToString <> "" Then
      sb.Append("LetterType,")
      sbv.Append(sInsUpdLetterType & ",")
    End If
    If sInsUpdEmailSent.ToString <> "" Then
      sb.Append("EmailSent,")
      sbv.Append(sInsUpdEmailSent & ",")
    End If
    If sInsUpdEmailSubject.ToString <> "" Then
      sb.Append("EmailSubject,")
      sbv.Append(sInsUpdEmailSubject & ",")
    End If
    If sInsUpdEmailContentType.ToString <> "" Then
      sb.Append("EmailContentType,")
      sbv.Append(sInsUpdEmailContentType & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(LogID) from [BookingsContactLog]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    LogID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [BookingsContactLog] Set ")
    If sInsUpdBookingID.ToString <> "" Then sb.Append("BookingID=" & sInsUpdBookingID & ",")
    If sInsUpdLogDate.ToString <> "" Then sb.Append("LogDate=" & sInsUpdLogDate & ",")
    If sInsUpdUser.ToString <> "" Then sb.Append("[User]=" & sInsUpdUser & ",")
    If sInsUpdLogEntry.ToString <> "" Then sb.Append("LogEntry=" & sInsUpdLogEntry & ",")
    If sInsUpdLogUpdates.ToString <> "" Then sb.Append("LogUpdates=" & sInsUpdLogUpdates & ",")
    If sInsUpdEmailAddress.ToString <> "" Then sb.Append("EmailAddress=" & sInsUpdEmailAddress & ",")
    If sInsUpdEmailType.ToString <> "" Then sb.Append("EmailType=" & sInsUpdEmailType & ",")
    If sInsUpdLetterType.ToString <> "" Then sb.Append("LetterType=" & sInsUpdLetterType & ",")
    If sInsUpdEmailSent.ToString <> "" Then sb.Append("EmailSent=" & sInsUpdEmailSent & ",")
    If sInsUpdEmailSubject.ToString <> "" Then sb.Append("EmailSubject=" & sInsUpdEmailSubject & ",")
    If sInsUpdEmailContentType.ToString <> "" Then sb.Append("EmailContentType=" & sInsUpdEmailContentType & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where LogID=" & sInsUpdLogID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [BookingsContactLog] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("LogID=" & sInsUpdLogID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TableErrorLog

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iID As Int32
  Private sInsUpdID As String
  Property ID_PK__Integer() As Int32
    Get
      Return iID
    End Get
    Set(ByVal Value As Int32)
      iID = Value
      sInsUpdID = oUtil.FixParam(iID, True)
    End Set
  End Property

  Private iErrNumber As Int32
  Private sInsUpdErrNumber As String
  Property ErrNumber__Integer() As Int32
    Get
      Return iErrNumber
    End Get
    Set(ByVal Value As Int32)
      iErrNumber = Value
      sInsUpdErrNumber = oUtil.FixParam(iErrNumber, True)
    End Set
  End Property

  Private sErrSource As String
  Private sInsUpdErrSource As String
  Property ErrSource__String() As String
    Get
      Return sErrSource
    End Get
    Set(ByVal Value As String)
      sErrSource = Value
      sInsUpdErrSource = oUtil.FixParam(sErrSource, True)
    End Set
  End Property

  Private sErrDescription As String
  Private sInsUpdErrDescription As String
  Property ErrDescription__String() As String
    Get
      Return sErrDescription
    End Get
    Set(ByVal Value As String)
      sErrDescription = Value
      sInsUpdErrDescription = oUtil.FixParam(sErrDescription, True)
    End Set
  End Property

  Private sErrPath As String
  Private sInsUpdErrPath As String
  Property ErrPath__String() As String
    Get
      Return sErrPath
    End Get
    Set(ByVal Value As String)
      sErrPath = Value
      sInsUpdErrPath = oUtil.FixParam(sErrPath, True)
    End Set
  End Property

  Private sErrComputer As String
  Private sInsUpdErrComputer As String
  Property ErrComputer__String() As String
    Get
      Return sErrComputer
    End Get
    Set(ByVal Value As String)
      sErrComputer = Value
      sInsUpdErrComputer = oUtil.FixParam(sErrComputer, True)
    End Set
  End Property

  Private sErrUsername As String
  Private sInsUpdErrUsername As String
  Property ErrUsername__String() As String
    Get
      Return sErrUsername
    End Get
    Set(ByVal Value As String)
      sErrUsername = Value
      sInsUpdErrUsername = oUtil.FixParam(sErrUsername, True)
    End Set
  End Property

  Private sErrEmailMessage As String
  Private sInsUpdErrEmailMessage As String
  Property ErrEmailMessage__String() As String
    Get
      Return sErrEmailMessage
    End Get
    Set(ByVal Value As String)
      sErrEmailMessage = Value
      sInsUpdErrEmailMessage = oUtil.FixParam(sErrEmailMessage, True)
    End Set
  End Property

  Private iErrWasEmailed As Int32
  Private sInsUpdErrWasEmailed As String
  Property ErrWasEmailed__Integer() As Int32
    Get
      Return iErrWasEmailed
    End Get
    Set(ByVal Value As Int32)
      iErrWasEmailed = Value
      sInsUpdErrWasEmailed = oUtil.FixParam(iErrWasEmailed, True)
    End Set
  End Property

  Private sErrTimeStamp As String
  Private sInsUpdErrTimeStamp As String
  Property ErrTimeStamp__Date() As String
    Get
      Return sErrTimeStamp
    End Get
    Set(ByVal Value As String)
      sErrTimeStamp = Value
      sInsUpdErrTimeStamp = oUtil.FixParam(sErrTimeStamp, True)
    End Set
  End Property

  Public Sub Clear()
    iID = 0
    sInsUpdID = ""
    iErrNumber = 0
    sInsUpdErrNumber = ""
    sErrSource = ""
    sInsUpdErrSource = ""
    sErrDescription = ""
    sInsUpdErrDescription = ""
    sErrPath = ""
    sInsUpdErrPath = ""
    sErrComputer = ""
    sInsUpdErrComputer = ""
    sErrUsername = ""
    sInsUpdErrUsername = ""
    sErrEmailMessage = ""
    sInsUpdErrEmailMessage = ""
    iErrWasEmailed = 0
    sInsUpdErrWasEmailed = ""
    sErrTimeStamp = ""
    sInsUpdErrTimeStamp = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdID.ToString = "'-12345'" Then sb.Append("ID,")
        If sInsUpdErrNumber.ToString = "'-12345'" Then sb.Append("ErrNumber,")
        If sInsUpdErrSource.ToString = "'Select'" Then sb.Append("ErrSource,")
        If sInsUpdErrDescription.ToString = "'Select'" Then sb.Append("ErrDescription,")
        If sInsUpdErrPath.ToString = "'Select'" Then sb.Append("ErrPath,")
        If sInsUpdErrComputer.ToString = "'Select'" Then sb.Append("ErrComputer,")
        If sInsUpdErrUsername.ToString = "'Select'" Then sb.Append("ErrUsername,")
        If sInsUpdErrEmailMessage.ToString = "'Select'" Then sb.Append("ErrEmailMessage,")
        If sInsUpdErrWasEmailed.ToString = "'-12345'" Then sb.Append("ErrWasEmailed,")
        If sInsUpdErrTimeStamp.ToString = "'Select'" Then sb.Append("ErrTimeStamp,")
      Else
        sb.Append("ID,")
        sb.Append("ErrNumber,")
        sb.Append("ErrSource,")
        sb.Append("ErrDescription,")
        sb.Append("ErrPath,")
        sb.Append("ErrComputer,")
        sb.Append("ErrUsername,")
        sb.Append("ErrEmailMessage,")
        sb.Append("ErrWasEmailed,")
        sb.Append("ErrTimeStamp,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [ErrorLog]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdID.ToString <> "") And (sInsUpdID <> "'-12345'") Then sbw.Append("ID=" & sInsUpdID & " and ")
      If (sInsUpdErrNumber.ToString <> "") And (sInsUpdErrNumber <> "'-12345'") Then sbw.Append("ErrNumber=" & sInsUpdErrNumber & " and ")
      If (sInsUpdErrSource.ToString <> "") And (sInsUpdErrSource <> "'Select'") Then sbw.Append("ErrSource=" & sInsUpdErrSource & " and ")
      If (sInsUpdErrDescription.ToString <> "") And (sInsUpdErrDescription <> "'Select'") Then sbw.Append("ErrDescription=" & sInsUpdErrDescription & " and ")
      If (sInsUpdErrPath.ToString <> "") And (sInsUpdErrPath <> "'Select'") Then sbw.Append("ErrPath=" & sInsUpdErrPath & " and ")
      If (sInsUpdErrComputer.ToString <> "") And (sInsUpdErrComputer <> "'Select'") Then sbw.Append("ErrComputer=" & sInsUpdErrComputer & " and ")
      If (sInsUpdErrUsername.ToString <> "") And (sInsUpdErrUsername <> "'Select'") Then sbw.Append("ErrUsername=" & sInsUpdErrUsername & " and ")
      If (sInsUpdErrEmailMessage.ToString <> "") And (sInsUpdErrEmailMessage <> "'Select'") Then sbw.Append("ErrEmailMessage=" & sInsUpdErrEmailMessage & " and ")
      If (sInsUpdErrWasEmailed.ToString <> "") And (sInsUpdErrWasEmailed <> "'-12345'") Then sbw.Append("ErrWasEmailed=" & sInsUpdErrWasEmailed & " and ")
      If (sInsUpdErrTimeStamp.ToString <> "") And (sInsUpdErrTimeStamp <> "'Select'") Then sbw.Append("ErrTimeStamp=" & sInsUpdErrTimeStamp & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      ID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("ID")), 0, CurrentRow.Item("ID").ToString)
      ErrNumber__Integer = IIf(IsDBNull(CurrentRow.Item("ErrNumber")), 0, CurrentRow.Item("ErrNumber"))
      ErrSource__String = IIf(IsDBNull(CurrentRow.Item("ErrSource")), "", CurrentRow.Item("ErrSource"))
      ErrDescription__String = IIf(IsDBNull(CurrentRow.Item("ErrDescription")), "", CurrentRow.Item("ErrDescription"))
      ErrPath__String = IIf(IsDBNull(CurrentRow.Item("ErrPath")), "", CurrentRow.Item("ErrPath"))
      ErrComputer__String = IIf(IsDBNull(CurrentRow.Item("ErrComputer")), "", CurrentRow.Item("ErrComputer"))
      ErrUsername__String = IIf(IsDBNull(CurrentRow.Item("ErrUsername")), "", CurrentRow.Item("ErrUsername"))
      ErrEmailMessage__String = IIf(IsDBNull(CurrentRow.Item("ErrEmailMessage")), "", CurrentRow.Item("ErrEmailMessage"))
      ErrWasEmailed__Integer = IIf(IsDBNull(CurrentRow.Item("ErrWasEmailed")), 0, CurrentRow.Item("ErrWasEmailed"))
      ErrTimeStamp__Date = IIf(IsDBNull(CurrentRow.Item("ErrTimeStamp")), "", CurrentRow.Item("ErrTimeStamp"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [ErrorLog](")
    If sInsUpdErrNumber.ToString <> "" Then
      sb.Append("ErrNumber,")
      sbv.Append(sInsUpdErrNumber & ",")
    End If
    If sInsUpdErrSource.ToString <> "" Then
      sb.Append("ErrSource,")
      sbv.Append(sInsUpdErrSource & ",")
    End If
    If sInsUpdErrDescription.ToString <> "" Then
      sb.Append("ErrDescription,")
      sbv.Append(sInsUpdErrDescription & ",")
    End If
    If sInsUpdErrPath.ToString <> "" Then
      sb.Append("ErrPath,")
      sbv.Append(sInsUpdErrPath & ",")
    End If
    If sInsUpdErrComputer.ToString <> "" Then
      sb.Append("ErrComputer,")
      sbv.Append(sInsUpdErrComputer & ",")
    End If
    If sInsUpdErrUsername.ToString <> "" Then
      sb.Append("ErrUsername,")
      sbv.Append(sInsUpdErrUsername & ",")
    End If
    If sInsUpdErrEmailMessage.ToString <> "" Then
      sb.Append("ErrEmailMessage,")
      sbv.Append(sInsUpdErrEmailMessage & ",")
    End If
    If sInsUpdErrWasEmailed.ToString <> "" Then
      sb.Append("ErrWasEmailed,")
      sbv.Append(sInsUpdErrWasEmailed & ",")
    End If
    If sInsUpdErrTimeStamp.ToString <> "" Then
      sb.Append("ErrTimeStamp,")
      sbv.Append(sInsUpdErrTimeStamp & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(ID) from [ErrorLog]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    ID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [ErrorLog] Set ")
    If sInsUpdErrNumber.ToString <> "" Then sb.Append("ErrNumber=" & sInsUpdErrNumber & ",")
    If sInsUpdErrSource.ToString <> "" Then sb.Append("ErrSource=" & sInsUpdErrSource & ",")
    If sInsUpdErrDescription.ToString <> "" Then sb.Append("ErrDescription=" & sInsUpdErrDescription & ",")
    If sInsUpdErrPath.ToString <> "" Then sb.Append("ErrPath=" & sInsUpdErrPath & ",")
    If sInsUpdErrComputer.ToString <> "" Then sb.Append("ErrComputer=" & sInsUpdErrComputer & ",")
    If sInsUpdErrUsername.ToString <> "" Then sb.Append("ErrUsername=" & sInsUpdErrUsername & ",")
    If sInsUpdErrEmailMessage.ToString <> "" Then sb.Append("ErrEmailMessage=" & sInsUpdErrEmailMessage & ",")
    If sInsUpdErrWasEmailed.ToString <> "" Then sb.Append("ErrWasEmailed=" & sInsUpdErrWasEmailed & ",")
    If sInsUpdErrTimeStamp.ToString <> "" Then sb.Append("ErrTimeStamp=" & sInsUpdErrTimeStamp & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where ID=" & sInsUpdID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [ErrorLog] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("ID=" & sInsUpdID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TableEvents

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private sEventName As String
  Private sInsUpdEventName As String
  Property EventName_PK__String() As String
    Get
      Return sEventName
    End Get
    Set(ByVal Value As String)
      sEventName = Value
      sInsUpdEventName = oUtil.FixParam(sEventName, True)
    End Set
  End Property

  Private sStartDate As String
  Private sInsUpdStartDate As String
  Property StartDate__Date() As String
    Get
      Return sStartDate
    End Get
    Set(ByVal Value As String)
      sStartDate = Value
      sInsUpdStartDate = oUtil.FixParam(sStartDate, True)
    End Set
  End Property

  Private sEndDate As String
  Private sInsUpdEndDate As String
  Property EndDate__Date() As String
    Get
      Return sEndDate
    End Get
    Set(ByVal Value As String)
      sEndDate = Value
      sInsUpdEndDate = oUtil.FixParam(sEndDate, True)
    End Set
  End Property

  Public Sub Clear()
    sEventName = ""
    sInsUpdEventName = ""
    sStartDate = ""
    sInsUpdStartDate = ""
    sEndDate = ""
    sInsUpdEndDate = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdEventName.ToString = "'Select'" Then sb.Append("EventName,")
        If sInsUpdStartDate.ToString = "'Select'" Then sb.Append("StartDate,")
        If sInsUpdEndDate.ToString = "'Select'" Then sb.Append("EndDate,")
      Else
        sb.Append("EventName,")
        sb.Append("StartDate,")
        sb.Append("EndDate,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [Events]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdEventName.ToString <> "") And (sInsUpdEventName <> "'Select'") Then sbw.Append("EventName=" & sInsUpdEventName & " and ")
      If (sInsUpdStartDate.ToString <> "") And (sInsUpdStartDate <> "'Select'") Then sbw.Append("StartDate=" & sInsUpdStartDate & " and ")
      If (sInsUpdEndDate.ToString <> "") And (sInsUpdEndDate <> "'Select'") Then sbw.Append("EndDate=" & sInsUpdEndDate & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      EventName_PK__String = IIf(IsDBNull(CurrentRow.Item("EventName")), "", CurrentRow.Item("EventName").ToString)
      StartDate__Date = IIf(IsDBNull(CurrentRow.Item("StartDate")), "", CurrentRow.Item("StartDate"))
      EndDate__Date = IIf(IsDBNull(CurrentRow.Item("EndDate")), "", CurrentRow.Item("EndDate"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [Events](")
    If sInsUpdEventName.ToString <> "" Then
      sb.Append("EventName,")
      sbv.Append(sInsUpdEventName & ",")
    End If
    If sInsUpdStartDate.ToString <> "" Then
      sb.Append("StartDate,")
      sbv.Append(sInsUpdStartDate & ",")
    End If
    If sInsUpdEndDate.ToString <> "" Then
      sb.Append("EndDate,")
      sbv.Append(sInsUpdEndDate & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [Events] Set ")
    If sInsUpdEventName.ToString <> "" Then sb.Append("EventName=" & sInsUpdEventName & ",")
    If sInsUpdStartDate.ToString <> "" Then sb.Append("StartDate=" & sInsUpdStartDate & ",")
    If sInsUpdEndDate.ToString <> "" Then sb.Append("EndDate=" & sInsUpdEndDate & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where EventName=" & sInsUpdEventName
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [Events] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("EventName=" & sInsUpdEventName)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TableHosts

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iHost_ID As Int32
  Private sInsUpdHost_ID As String
  Property Host_ID_PK__Integer() As Int32
    Get
      Return iHost_ID
    End Get
    Set(ByVal Value As Int32)
      iHost_ID = Value
      sInsUpdHost_ID = oUtil.FixParam(iHost_ID, True)
    End Set
  End Property

  Private sHostName As String
  Private sInsUpdHostName As String
  Property HostName__String() As String
    Get
      Return sHostName
    End Get
    Set(ByVal Value As String)
      sHostName = Value
      sInsUpdHostName = oUtil.FixParam(sHostName, True)
    End Set
  End Property

  Private sHostLastName As String
  Private sInsUpdHostLastName As String
  Property HostLastName__String() As String
    Get
      Return sHostLastName
    End Get
    Set(ByVal Value As String)
      sHostLastName = Value
      sInsUpdHostLastName = oUtil.FixParam(sHostLastName, True)
    End Set
  End Property

  Private sAddress As String
  Private sInsUpdAddress As String
  Property Address__String() As String
    Get
      Return sAddress
    End Get
    Set(ByVal Value As String)
      sAddress = Value
      sInsUpdAddress = oUtil.FixParam(sAddress, True)
    End Set
  End Property

  Private sAddress2 As String
  Private sInsUpdAddress2 As String
  Property Address2__String() As String
    Get
      Return sAddress2
    End Get
    Set(ByVal Value As String)
      sAddress2 = Value
      sInsUpdAddress2 = oUtil.FixParam(sAddress2, True)
    End Set
  End Property

  Private sCity As String
  Private sInsUpdCity As String
  Property City__String() As String
    Get
      Return sCity
    End Get
    Set(ByVal Value As String)
      sCity = Value
      sInsUpdCity = oUtil.FixParam(sCity, True)
    End Set
  End Property

  Private sZip As String
  Private sInsUpdZip As String
  Property Zip__String() As String
    Get
      Return sZip
    End Get
    Set(ByVal Value As String)
      sZip = Value
      sInsUpdZip = oUtil.FixParam(sZip, True)
    End Set
  End Property

  Private sHomePhone As String
  Private sInsUpdHomePhone As String
  Property HomePhone__String() As String
    Get
      Return sHomePhone
    End Get
    Set(ByVal Value As String)
      sHomePhone = Value
      sInsUpdHomePhone = oUtil.FixParam(sHomePhone, True)
    End Set
  End Property

  Private sWorkPhone As String
  Private sInsUpdWorkPhone As String
  Property WorkPhone__String() As String
    Get
      Return sWorkPhone
    End Get
    Set(ByVal Value As String)
      sWorkPhone = Value
      sInsUpdWorkPhone = oUtil.FixParam(sWorkPhone, True)
    End Set
  End Property

  Private sCellPhone As String
  Private sInsUpdCellPhone As String
  Property CellPhone__String() As String
    Get
      Return sCellPhone
    End Get
    Set(ByVal Value As String)
      sCellPhone = Value
      sInsUpdCellPhone = oUtil.FixParam(sCellPhone, True)
    End Set
  End Property

  Private sEmail As String
  Private sInsUpdEmail As String
  Property Email__String() As String
    Get
      Return sEmail
    End Get
    Set(ByVal Value As String)
      sEmail = Value
      sInsUpdEmail = oUtil.FixParam(sEmail, True)
    End Set
  End Property

  Private sHostNotes As String
  Private sInsUpdHostNotes As String
  Property HostNotes__String() As String
    Get
      Return sHostNotes
    End Get
    Set(ByVal Value As String)
      sHostNotes = Value
      sInsUpdHostNotes = oUtil.FixParam(sHostNotes, True)
    End Set
  End Property

  Private sCellPhone2 As String
  Private sInsUpdCellPhone2 As String
  Property CellPhone2__String() As String
    Get
      Return sCellPhone2
    End Get
    Set(ByVal Value As String)
      sCellPhone2 = Value
      sInsUpdCellPhone2 = oUtil.FixParam(sCellPhone2, True)
    End Set
  End Property

  Private sEmail2 As String
  Private sInsUpdEmail2 As String
  Property Email2__String() As String
    Get
      Return sEmail2
    End Get
    Set(ByVal Value As String)
      sEmail2 = Value
      sInsUpdEmail2 = oUtil.FixParam(sEmail2, True)
    End Set
  End Property

  Private sCheckName As String
  Private sInsUpdCheckName As String
  Property CheckName__String() As String
    Get
      Return sCheckName
    End Get
    Set(ByVal Value As String)
      sCheckName = Value
      sInsUpdCheckName = oUtil.FixParam(sCheckName, True)
    End Set
  End Property

  Private sStatus As String
  Private sInsUpdStatus As String
  Property Status__String() As String
    Get
      Return sStatus
    End Get
    Set(ByVal Value As String)
      sStatus = Value
      sInsUpdStatus = oUtil.FixParam(sStatus, True)
    End Set
  End Property

  Private sQBListID As String
  Private sInsUpdQBListID As String
  Property QBListID__String() As String
    Get
      Return sQBListID
    End Get
    Set(ByVal Value As String)
      sQBListID = Value
      sInsUpdQBListID = oUtil.FixParam(sQBListID, True)
    End Set
  End Property

  Private sContractYears As String
  Private sInsUpdContractYears As String
  Property ContractYears__String() As String
    Get
      Return sContractYears
    End Get
    Set(ByVal Value As String)
      sContractYears = Value
      sInsUpdContractYears = oUtil.FixParam(sContractYears, True)
    End Set
  End Property

  Public Sub Clear()
    iHost_ID = 0
    sInsUpdHost_ID = ""
    sHostName = ""
    sInsUpdHostName = ""
    sHostLastName = ""
    sInsUpdHostLastName = ""
    sAddress = ""
    sInsUpdAddress = ""
    sAddress2 = ""
    sInsUpdAddress2 = ""
    sCity = ""
    sInsUpdCity = ""
    sZip = ""
    sInsUpdZip = ""
    sHomePhone = ""
    sInsUpdHomePhone = ""
    sWorkPhone = ""
    sInsUpdWorkPhone = ""
    sCellPhone = ""
    sInsUpdCellPhone = ""
    sEmail = ""
    sInsUpdEmail = ""
    sHostNotes = ""
    sInsUpdHostNotes = ""
    sCellPhone2 = ""
    sInsUpdCellPhone2 = ""
    sEmail2 = ""
    sInsUpdEmail2 = ""
    sCheckName = ""
    sInsUpdCheckName = ""
    sStatus = ""
    sInsUpdStatus = ""
    sQBListID = ""
    sInsUpdQBListID = ""
    sContractYears = ""
    sInsUpdContractYears = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdHost_ID.ToString = "'-12345'" Then sb.Append("Host_ID,")
        If sInsUpdHostName.ToString = "'Select'" Then sb.Append("HostName,")
        If sInsUpdHostLastName.ToString = "'Select'" Then sb.Append("HostLastName,")
        If sInsUpdAddress.ToString = "'Select'" Then sb.Append("Address,")
        If sInsUpdAddress2.ToString = "'Select'" Then sb.Append("Address2,")
        If sInsUpdCity.ToString = "'Select'" Then sb.Append("City,")
        If sInsUpdZip.ToString = "'Select'" Then sb.Append("Zip,")
        If sInsUpdHomePhone.ToString = "'Select'" Then sb.Append("HomePhone,")
        If sInsUpdWorkPhone.ToString = "'Select'" Then sb.Append("WorkPhone,")
        If sInsUpdCellPhone.ToString = "'Select'" Then sb.Append("CellPhone,")
        If sInsUpdEmail.ToString = "'Select'" Then sb.Append("Email,")
        If sInsUpdHostNotes.ToString = "'Select'" Then sb.Append("HostNotes,")
        If sInsUpdCellPhone2.ToString = "'Select'" Then sb.Append("CellPhone2,")
        If sInsUpdEmail2.ToString = "'Select'" Then sb.Append("Email2,")
        If sInsUpdCheckName.ToString = "'Select'" Then sb.Append("CheckName,")
        If sInsUpdStatus.ToString = "'Select'" Then sb.Append("Status,")
        If sInsUpdQBListID.ToString = "'Select'" Then sb.Append("QBListID,")
        If sInsUpdContractYears.ToString = "'Select'" Then sb.Append("ContractYears,")
      Else
        sb.Append("Host_ID,")
        sb.Append("HostName,")
        sb.Append("HostLastName,")
        sb.Append("Address,")
        sb.Append("Address2,")
        sb.Append("City,")
        sb.Append("Zip,")
        sb.Append("HomePhone,")
        sb.Append("WorkPhone,")
        sb.Append("CellPhone,")
        sb.Append("Email,")
        sb.Append("HostNotes,")
        sb.Append("CellPhone2,")
        sb.Append("Email2,")
        sb.Append("CheckName,")
        sb.Append("Status,")
        sb.Append("QBListID,")
        sb.Append("ContractYears,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [Hosts]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdHost_ID.ToString <> "") And (sInsUpdHost_ID <> "'-12345'") Then sbw.Append("Host_ID=" & sInsUpdHost_ID & " and ")
      If (sInsUpdHostName.ToString <> "") And (sInsUpdHostName <> "'Select'") Then sbw.Append("HostName=" & sInsUpdHostName & " and ")
      If (sInsUpdHostLastName.ToString <> "") And (sInsUpdHostLastName <> "'Select'") Then sbw.Append("HostLastName=" & sInsUpdHostLastName & " and ")
      If (sInsUpdAddress.ToString <> "") And (sInsUpdAddress <> "'Select'") Then sbw.Append("Address=" & sInsUpdAddress & " and ")
      If (sInsUpdAddress2.ToString <> "") And (sInsUpdAddress2 <> "'Select'") Then sbw.Append("Address2=" & sInsUpdAddress2 & " and ")
      If (sInsUpdCity.ToString <> "") And (sInsUpdCity <> "'Select'") Then sbw.Append("City=" & sInsUpdCity & " and ")
      If (sInsUpdZip.ToString <> "") And (sInsUpdZip <> "'Select'") Then sbw.Append("Zip=" & sInsUpdZip & " and ")
      If (sInsUpdHomePhone.ToString <> "") And (sInsUpdHomePhone <> "'Select'") Then sbw.Append("HomePhone=" & sInsUpdHomePhone & " and ")
      If (sInsUpdWorkPhone.ToString <> "") And (sInsUpdWorkPhone <> "'Select'") Then sbw.Append("WorkPhone=" & sInsUpdWorkPhone & " and ")
      If (sInsUpdCellPhone.ToString <> "") And (sInsUpdCellPhone <> "'Select'") Then sbw.Append("CellPhone=" & sInsUpdCellPhone & " and ")
      If (sInsUpdEmail.ToString <> "") And (sInsUpdEmail <> "'Select'") Then sbw.Append("Email=" & sInsUpdEmail & " and ")
      If (sInsUpdHostNotes.ToString <> "") And (sInsUpdHostNotes <> "'Select'") Then sbw.Append("HostNotes=" & sInsUpdHostNotes & " and ")
      If (sInsUpdCellPhone2.ToString <> "") And (sInsUpdCellPhone2 <> "'Select'") Then sbw.Append("CellPhone2=" & sInsUpdCellPhone2 & " and ")
      If (sInsUpdEmail2.ToString <> "") And (sInsUpdEmail2 <> "'Select'") Then sbw.Append("Email2=" & sInsUpdEmail2 & " and ")
      If (sInsUpdCheckName.ToString <> "") And (sInsUpdCheckName <> "'Select'") Then sbw.Append("CheckName=" & sInsUpdCheckName & " and ")
      If (sInsUpdStatus.ToString <> "") And (sInsUpdStatus <> "'Select'") Then sbw.Append("Status=" & sInsUpdStatus & " and ")
      If (sInsUpdQBListID.ToString <> "") And (sInsUpdQBListID <> "'Select'") Then sbw.Append("QBListID=" & sInsUpdQBListID & " and ")
      If (sInsUpdContractYears.ToString <> "") And (sInsUpdContractYears <> "'Select'") Then sbw.Append("ContractYears=" & sInsUpdContractYears & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      Host_ID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("Host_ID")), 0, CurrentRow.Item("Host_ID").ToString)
      HostName__String = IIf(IsDBNull(CurrentRow.Item("HostName")), "", CurrentRow.Item("HostName"))
      HostLastName__String = IIf(IsDBNull(CurrentRow.Item("HostLastName")), "", CurrentRow.Item("HostLastName"))
      Address__String = IIf(IsDBNull(CurrentRow.Item("Address")), "", CurrentRow.Item("Address"))
      Address2__String = IIf(IsDBNull(CurrentRow.Item("Address2")), "", CurrentRow.Item("Address2"))
      City__String = IIf(IsDBNull(CurrentRow.Item("City")), "", CurrentRow.Item("City"))
      Zip__String = IIf(IsDBNull(CurrentRow.Item("Zip")), "", CurrentRow.Item("Zip"))
      HomePhone__String = IIf(IsDBNull(CurrentRow.Item("HomePhone")), "", CurrentRow.Item("HomePhone"))
      WorkPhone__String = IIf(IsDBNull(CurrentRow.Item("WorkPhone")), "", CurrentRow.Item("WorkPhone"))
      CellPhone__String = IIf(IsDBNull(CurrentRow.Item("CellPhone")), "", CurrentRow.Item("CellPhone"))
      Email__String = IIf(IsDBNull(CurrentRow.Item("Email")), "", CurrentRow.Item("Email"))
      HostNotes__String = IIf(IsDBNull(CurrentRow.Item("HostNotes")), "", CurrentRow.Item("HostNotes"))
      CellPhone2__String = IIf(IsDBNull(CurrentRow.Item("CellPhone2")), "", CurrentRow.Item("CellPhone2"))
      Email2__String = IIf(IsDBNull(CurrentRow.Item("Email2")), "", CurrentRow.Item("Email2"))
      CheckName__String = IIf(IsDBNull(CurrentRow.Item("CheckName")), "", CurrentRow.Item("CheckName"))
      Status__String = IIf(IsDBNull(CurrentRow.Item("Status")), "", CurrentRow.Item("Status"))
      QBListID__String = IIf(IsDBNull(CurrentRow.Item("QBListID")), "", CurrentRow.Item("QBListID"))
      ContractYears__String = IIf(IsDBNull(CurrentRow.Item("ContractYears")), "", CurrentRow.Item("ContractYears"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [Hosts](")
    If sInsUpdHostName.ToString <> "" Then
      sb.Append("HostName,")
      sbv.Append(sInsUpdHostName & ",")
    End If
    If sInsUpdHostLastName.ToString <> "" Then
      sb.Append("HostLastName,")
      sbv.Append(sInsUpdHostLastName & ",")
    End If
    If sInsUpdAddress.ToString <> "" Then
      sb.Append("Address,")
      sbv.Append(sInsUpdAddress & ",")
    End If
    If sInsUpdAddress2.ToString <> "" Then
      sb.Append("Address2,")
      sbv.Append(sInsUpdAddress2 & ",")
    End If
    If sInsUpdCity.ToString <> "" Then
      sb.Append("City,")
      sbv.Append(sInsUpdCity & ",")
    End If
    If sInsUpdZip.ToString <> "" Then
      sb.Append("Zip,")
      sbv.Append(sInsUpdZip & ",")
    End If
    If sInsUpdHomePhone.ToString <> "" Then
      sb.Append("HomePhone,")
      sbv.Append(sInsUpdHomePhone & ",")
    End If
    If sInsUpdWorkPhone.ToString <> "" Then
      sb.Append("WorkPhone,")
      sbv.Append(sInsUpdWorkPhone & ",")
    End If
    If sInsUpdCellPhone.ToString <> "" Then
      sb.Append("CellPhone,")
      sbv.Append(sInsUpdCellPhone & ",")
    End If
    If sInsUpdEmail.ToString <> "" Then
      sb.Append("Email,")
      sbv.Append(sInsUpdEmail & ",")
    End If
    If sInsUpdHostNotes.ToString <> "" Then
      sb.Append("HostNotes,")
      sbv.Append(sInsUpdHostNotes & ",")
    End If
    If sInsUpdCellPhone2.ToString <> "" Then
      sb.Append("CellPhone2,")
      sbv.Append(sInsUpdCellPhone2 & ",")
    End If
    If sInsUpdEmail2.ToString <> "" Then
      sb.Append("Email2,")
      sbv.Append(sInsUpdEmail2 & ",")
    End If
    If sInsUpdCheckName.ToString <> "" Then
      sb.Append("CheckName,")
      sbv.Append(sInsUpdCheckName & ",")
    End If
    If sInsUpdStatus.ToString <> "" Then
      sb.Append("Status,")
      sbv.Append(sInsUpdStatus & ",")
    End If
    If sInsUpdQBListID.ToString <> "" Then
      sb.Append("QBListID,")
      sbv.Append(sInsUpdQBListID & ",")
    End If
    If sInsUpdContractYears.ToString <> "" Then
      sb.Append("ContractYears,")
      sbv.Append(sInsUpdContractYears & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(Host_ID) from [Hosts]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    Host_ID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [Hosts] Set ")
    If sInsUpdHostName.ToString <> "" Then sb.Append("HostName=" & sInsUpdHostName & ",")
    If sInsUpdHostLastName.ToString <> "" Then sb.Append("HostLastName=" & sInsUpdHostLastName & ",")
    If sInsUpdAddress.ToString <> "" Then sb.Append("Address=" & sInsUpdAddress & ",")
    If sInsUpdAddress2.ToString <> "" Then sb.Append("Address2=" & sInsUpdAddress2 & ",")
    If sInsUpdCity.ToString <> "" Then sb.Append("City=" & sInsUpdCity & ",")
    If sInsUpdZip.ToString <> "" Then sb.Append("Zip=" & sInsUpdZip & ",")
    If sInsUpdHomePhone.ToString <> "" Then sb.Append("HomePhone=" & sInsUpdHomePhone & ",")
    If sInsUpdWorkPhone.ToString <> "" Then sb.Append("WorkPhone=" & sInsUpdWorkPhone & ",")
    If sInsUpdCellPhone.ToString <> "" Then sb.Append("CellPhone=" & sInsUpdCellPhone & ",")
    If sInsUpdEmail.ToString <> "" Then sb.Append("Email=" & sInsUpdEmail & ",")
    If sInsUpdHostNotes.ToString <> "" Then sb.Append("HostNotes=" & sInsUpdHostNotes & ",")
    If sInsUpdCellPhone2.ToString <> "" Then sb.Append("CellPhone2=" & sInsUpdCellPhone2 & ",")
    If sInsUpdEmail2.ToString <> "" Then sb.Append("Email2=" & sInsUpdEmail2 & ",")
    If sInsUpdCheckName.ToString <> "" Then sb.Append("CheckName=" & sInsUpdCheckName & ",")
    If sInsUpdStatus.ToString <> "" Then sb.Append("Status=" & sInsUpdStatus & ",")
    If sInsUpdQBListID.ToString <> "" Then sb.Append("QBListID=" & sInsUpdQBListID & ",")
    If sInsUpdContractYears.ToString <> "" Then sb.Append("ContractYears=" & sInsUpdContractYears & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where Host_ID=" & sInsUpdHost_ID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [Hosts] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("Host_ID=" & sInsUpdHost_ID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TableGuestsContactLog

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iLogID As Int32
  Private sInsUpdLogID As String
  Property LogID_PK__Integer() As Int32
    Get
      Return iLogID
    End Get
    Set(ByVal Value As Int32)
      iLogID = Value
      sInsUpdLogID = oUtil.FixParam(iLogID, True)
    End Set
  End Property

  Private iGuestID As Int32
  Private sInsUpdGuestID As String
  Property GuestID__Integer() As Int32
    Get
      Return iGuestID
    End Get
    Set(ByVal Value As Int32)
      iGuestID = Value
      sInsUpdGuestID = oUtil.FixParam(iGuestID, True)
    End Set
  End Property

  Private sLogDate As String
  Private sInsUpdLogDate As String
  Property LogDate__Date() As String
    Get
      Return sLogDate
    End Get
    Set(ByVal Value As String)
      sLogDate = Value
      sInsUpdLogDate = oUtil.FixParam(sLogDate, True)
    End Set
  End Property

  Private sUser As String
  Private sInsUpdUser As String
  Property User__String() As String
    Get
      Return sUser
    End Get
    Set(ByVal Value As String)
      sUser = Value
      sInsUpdUser = oUtil.FixParam(sUser, True)
    End Set
  End Property

  Private sLogEntry As String
  Private sInsUpdLogEntry As String
  Property LogEntry__String() As String
    Get
      Return sLogEntry
    End Get
    Set(ByVal Value As String)
      sLogEntry = Value
      sInsUpdLogEntry = oUtil.FixParam(sLogEntry, True)
    End Set
  End Property

  Private sLogUpdates As String
  Private sInsUpdLogUpdates As String
  Property LogUpdates__String() As String
    Get
      Return sLogUpdates
    End Get
    Set(ByVal Value As String)
      sLogUpdates = Value
      sInsUpdLogUpdates = oUtil.FixParam(sLogUpdates, True)
    End Set
  End Property

  Public Sub Clear()
    iLogID = 0
    sInsUpdLogID = ""
    iGuestID = 0
    sInsUpdGuestID = ""
    sLogDate = ""
    sInsUpdLogDate = ""
    sUser = ""
    sInsUpdUser = ""
    sLogEntry = ""
    sInsUpdLogEntry = ""
    sLogUpdates = ""
    sInsUpdLogUpdates = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdLogID.ToString = "'-12345'" Then sb.Append("LogID,")
        If sInsUpdGuestID.ToString = "'-12345'" Then sb.Append("GuestID,")
        If sInsUpdLogDate.ToString = "'Select'" Then sb.Append("LogDate,")
        If sInsUpdUser.ToString = "'Select'" Then sb.Append("[User],")
        If sInsUpdLogEntry.ToString = "'Select'" Then sb.Append("LogEntry,")
        If sInsUpdLogUpdates.ToString = "'Select'" Then sb.Append("LogUpdates,")
      Else
        sb.Append("LogID,")
        sb.Append("GuestID,")
        sb.Append("LogDate,")
        sb.Append("[User],")
        sb.Append("LogEntry,")
        sb.Append("LogUpdates,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [GuestsContactLog]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdLogID.ToString <> "") And (sInsUpdLogID <> "'-12345'") Then sbw.Append("LogID=" & sInsUpdLogID & " and ")
      If (sInsUpdGuestID.ToString <> "") And (sInsUpdGuestID <> "'-12345'") Then sbw.Append("GuestID=" & sInsUpdGuestID & " and ")
      If (sInsUpdLogDate.ToString <> "") And (sInsUpdLogDate <> "'Select'") Then sbw.Append("LogDate=" & sInsUpdLogDate & " and ")
      If (sInsUpdUser.ToString <> "") And (sInsUpdUser <> "'Select'") Then sbw.Append("[User]=" & sInsUpdUser & " and ")
      If (sInsUpdLogEntry.ToString <> "") And (sInsUpdLogEntry <> "'Select'") Then sbw.Append("LogEntry=" & sInsUpdLogEntry & " and ")
      If (sInsUpdLogUpdates.ToString <> "") And (sInsUpdLogUpdates <> "'Select'") Then sbw.Append("LogUpdates=" & sInsUpdLogUpdates & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      LogID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("LogID")), 0, CurrentRow.Item("LogID").ToString)
      GuestID__Integer = IIf(IsDBNull(CurrentRow.Item("GuestID")), 0, CurrentRow.Item("GuestID"))
      LogDate__Date = IIf(IsDBNull(CurrentRow.Item("LogDate")), "", CurrentRow.Item("LogDate"))
      User__String = IIf(IsDBNull(CurrentRow.Item("User")), "", CurrentRow.Item("User"))
      LogEntry__String = IIf(IsDBNull(CurrentRow.Item("LogEntry")), "", CurrentRow.Item("LogEntry"))
      LogUpdates__String = IIf(IsDBNull(CurrentRow.Item("LogUpdates")), "", CurrentRow.Item("LogUpdates"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [GuestsContactLog](")
    If sInsUpdGuestID.ToString <> "" Then
      sb.Append("GuestID,")
      sbv.Append(sInsUpdGuestID & ",")
    End If
    If sInsUpdLogDate.ToString <> "" Then
      sb.Append("LogDate,")
      sbv.Append(sInsUpdLogDate & ",")
    End If
    If sInsUpdUser.ToString <> "" Then
      sb.Append("[User],")
      sbv.Append(sInsUpdUser & ",")
    End If
    If sInsUpdLogEntry.ToString <> "" Then
      sb.Append("LogEntry,")
      sbv.Append(sInsUpdLogEntry & ",")
    End If
    If sInsUpdLogUpdates.ToString <> "" Then
      sb.Append("LogUpdates,")
      sbv.Append(sInsUpdLogUpdates & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(LogID) from [GuestsContactLog]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    LogID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [GuestsContactLog] Set ")
    If sInsUpdGuestID.ToString <> "" Then sb.Append("GuestID=" & sInsUpdGuestID & ",")
    If sInsUpdLogDate.ToString <> "" Then sb.Append("LogDate=" & sInsUpdLogDate & ",")
    If sInsUpdUser.ToString <> "" Then sb.Append("[User]=" & sInsUpdUser & ",")
    If sInsUpdLogEntry.ToString <> "" Then sb.Append("LogEntry=" & sInsUpdLogEntry & ",")
    If sInsUpdLogUpdates.ToString <> "" Then sb.Append("LogUpdates=" & sInsUpdLogUpdates & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where LogID=" & sInsUpdLogID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [GuestsContactLog] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("LogID=" & sInsUpdLogID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TableWebErrorLog

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iidError As Int32
  Private sInsUpdidError As String
  Property idError_PK__Integer() As Int32
    Get
      Return iidError
    End Get
    Set(ByVal Value As Int32)
      iidError = Value
      sInsUpdidError = oUtil.FixParam(iidError, True)
    End Set
  End Property

  Private sdtErrorDate As String
  Private sInsUpddtErrorDate As String
  Property dtErrorDate__Date() As String
    Get
      Return sdtErrorDate
    End Get
    Set(ByVal Value As String)
      sdtErrorDate = Value
      sInsUpddtErrorDate = oUtil.FixParam(sdtErrorDate, True)
    End Set
  End Property

  Private icErrorNumber As Int32
  Private sInsUpdcErrorNumber As String
  Property cErrorNumber__Integer() As Int32
    Get
      Return icErrorNumber
    End Get
    Set(ByVal Value As Int32)
      icErrorNumber = Value
      sInsUpdcErrorNumber = oUtil.FixParam(icErrorNumber, True)
    End Set
  End Property

  Private scErrorSource As String
  Private sInsUpdcErrorSource As String
  Property cErrorSource__String() As String
    Get
      Return scErrorSource
    End Get
    Set(ByVal Value As String)
      scErrorSource = Value
      sInsUpdcErrorSource = oUtil.FixParam(scErrorSource, True)
    End Set
  End Property

  Private scErrorMessage As String
  Private sInsUpdcErrorMessage As String
  Property cErrorMessage__String() As String
    Get
      Return scErrorMessage
    End Get
    Set(ByVal Value As String)
      scErrorMessage = Value
      sInsUpdcErrorMessage = oUtil.FixParam(scErrorMessage, True)
    End Set
  End Property

  Private scErrorTargetSite As String
  Private sInsUpdcErrorTargetSite As String
  Property cErrorTargetSite__String() As String
    Get
      Return scErrorTargetSite
    End Get
    Set(ByVal Value As String)
      scErrorTargetSite = Value
      sInsUpdcErrorTargetSite = oUtil.FixParam(scErrorTargetSite, True)
    End Set
  End Property

  Private scErrorStackTrace As String
  Private sInsUpdcErrorStackTrace As String
  Property cErrorStackTrace__String() As String
    Get
      Return scErrorStackTrace
    End Get
    Set(ByVal Value As String)
      scErrorStackTrace = Value
      sInsUpdcErrorStackTrace = oUtil.FixParam(scErrorStackTrace, True)
    End Set
  End Property

  Private scErrorHTML As String
  Private sInsUpdcErrorHTML As String
  Property cErrorHTML__String() As String
    Get
      Return scErrorHTML
    End Get
    Set(ByVal Value As String)
      scErrorHTML = Value
      sInsUpdcErrorHTML = oUtil.FixParam(scErrorHTML, True)
    End Set
  End Property

  Private scErrorUserInput As String
  Private sInsUpdcErrorUserInput As String
  Property cErrorUserInput__String() As String
    Get
      Return scErrorUserInput
    End Get
    Set(ByVal Value As String)
      scErrorUserInput = Value
      sInsUpdcErrorUserInput = oUtil.FixParam(scErrorUserInput, True)
    End Set
  End Property

  Public Sub Clear()
    iidError = 0
    sInsUpdidError = ""
    sdtErrorDate = ""
    sInsUpddtErrorDate = ""
    icErrorNumber = 0
    sInsUpdcErrorNumber = ""
    scErrorSource = ""
    sInsUpdcErrorSource = ""
    scErrorMessage = ""
    sInsUpdcErrorMessage = ""
    scErrorTargetSite = ""
    sInsUpdcErrorTargetSite = ""
    scErrorStackTrace = ""
    sInsUpdcErrorStackTrace = ""
    scErrorHTML = ""
    sInsUpdcErrorHTML = ""
    scErrorUserInput = ""
    sInsUpdcErrorUserInput = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdidError.ToString = "'-12345'" Then sb.Append("idError,")
        If sInsUpddtErrorDate.ToString = "'Select'" Then sb.Append("dtErrorDate,")
        If sInsUpdcErrorNumber.ToString = "'-12345'" Then sb.Append("cErrorNumber,")
        If sInsUpdcErrorSource.ToString = "'Select'" Then sb.Append("cErrorSource,")
        If sInsUpdcErrorMessage.ToString = "'Select'" Then sb.Append("cErrorMessage,")
        If sInsUpdcErrorTargetSite.ToString = "'Select'" Then sb.Append("cErrorTargetSite,")
        If sInsUpdcErrorStackTrace.ToString = "'Select'" Then sb.Append("cErrorStackTrace,")
        If sInsUpdcErrorHTML.ToString = "'Select'" Then sb.Append("cErrorHTML,")
        If sInsUpdcErrorUserInput.ToString = "'Select'" Then sb.Append("cErrorUserInput,")
      Else
        sb.Append("idError,")
        sb.Append("dtErrorDate,")
        sb.Append("cErrorNumber,")
        sb.Append("cErrorSource,")
        sb.Append("cErrorMessage,")
        sb.Append("cErrorTargetSite,")
        sb.Append("cErrorStackTrace,")
        sb.Append("cErrorHTML,")
        sb.Append("cErrorUserInput,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [WebErrorLog]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdidError.ToString <> "") And (sInsUpdidError <> "'-12345'") Then sbw.Append("idError=" & sInsUpdidError & " and ")
      If (sInsUpddtErrorDate.ToString <> "") And (sInsUpddtErrorDate <> "'Select'") Then sbw.Append("dtErrorDate=" & sInsUpddtErrorDate & " and ")
      If (sInsUpdcErrorNumber.ToString <> "") And (sInsUpdcErrorNumber <> "'-12345'") Then sbw.Append("cErrorNumber=" & sInsUpdcErrorNumber & " and ")
      If (sInsUpdcErrorSource.ToString <> "") And (sInsUpdcErrorSource <> "'Select'") Then sbw.Append("cErrorSource=" & sInsUpdcErrorSource & " and ")
      If (sInsUpdcErrorMessage.ToString <> "") And (sInsUpdcErrorMessage <> "'Select'") Then sbw.Append("cErrorMessage=" & sInsUpdcErrorMessage & " and ")
      If (sInsUpdcErrorTargetSite.ToString <> "") And (sInsUpdcErrorTargetSite <> "'Select'") Then sbw.Append("cErrorTargetSite=" & sInsUpdcErrorTargetSite & " and ")
      If (sInsUpdcErrorStackTrace.ToString <> "") And (sInsUpdcErrorStackTrace <> "'Select'") Then sbw.Append("cErrorStackTrace=" & sInsUpdcErrorStackTrace & " and ")
      If (sInsUpdcErrorHTML.ToString <> "") And (sInsUpdcErrorHTML <> "'Select'") Then sbw.Append("cErrorHTML=" & sInsUpdcErrorHTML & " and ")
      If (sInsUpdcErrorUserInput.ToString <> "") And (sInsUpdcErrorUserInput <> "'Select'") Then sbw.Append("cErrorUserInput=" & sInsUpdcErrorUserInput & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      idError_PK__Integer = IIf(IsDBNull(CurrentRow.Item("idError")), 0, CurrentRow.Item("idError").ToString)
      dtErrorDate__Date = IIf(IsDBNull(CurrentRow.Item("dtErrorDate")), "", CurrentRow.Item("dtErrorDate"))
      cErrorNumber__Integer = IIf(IsDBNull(CurrentRow.Item("cErrorNumber")), 0, CurrentRow.Item("cErrorNumber"))
      cErrorSource__String = IIf(IsDBNull(CurrentRow.Item("cErrorSource")), "", CurrentRow.Item("cErrorSource"))
      cErrorMessage__String = IIf(IsDBNull(CurrentRow.Item("cErrorMessage")), "", CurrentRow.Item("cErrorMessage"))
      cErrorTargetSite__String = IIf(IsDBNull(CurrentRow.Item("cErrorTargetSite")), "", CurrentRow.Item("cErrorTargetSite"))
      cErrorStackTrace__String = IIf(IsDBNull(CurrentRow.Item("cErrorStackTrace")), "", CurrentRow.Item("cErrorStackTrace"))
      cErrorHTML__String = IIf(IsDBNull(CurrentRow.Item("cErrorHTML")), "", CurrentRow.Item("cErrorHTML"))
      cErrorUserInput__String = IIf(IsDBNull(CurrentRow.Item("cErrorUserInput")), "", CurrentRow.Item("cErrorUserInput"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [WebErrorLog](")
    If sInsUpddtErrorDate.ToString <> "" Then
      sb.Append("dtErrorDate,")
      sbv.Append(sInsUpddtErrorDate & ",")
    End If
    If sInsUpdcErrorNumber.ToString <> "" Then
      sb.Append("cErrorNumber,")
      sbv.Append(sInsUpdcErrorNumber & ",")
    End If
    If sInsUpdcErrorSource.ToString <> "" Then
      sb.Append("cErrorSource,")
      sbv.Append(sInsUpdcErrorSource & ",")
    End If
    If sInsUpdcErrorMessage.ToString <> "" Then
      sb.Append("cErrorMessage,")
      sbv.Append(sInsUpdcErrorMessage & ",")
    End If
    If sInsUpdcErrorTargetSite.ToString <> "" Then
      sb.Append("cErrorTargetSite,")
      sbv.Append(sInsUpdcErrorTargetSite & ",")
    End If
    If sInsUpdcErrorStackTrace.ToString <> "" Then
      sb.Append("cErrorStackTrace,")
      sbv.Append(sInsUpdcErrorStackTrace & ",")
    End If
    If sInsUpdcErrorHTML.ToString <> "" Then
      sb.Append("cErrorHTML,")
      sbv.Append(sInsUpdcErrorHTML & ",")
    End If
    If sInsUpdcErrorUserInput.ToString <> "" Then
      sb.Append("cErrorUserInput,")
      sbv.Append(sInsUpdcErrorUserInput & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(idError) from [WebErrorLog]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    idError_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [WebErrorLog] Set ")
    If sInsUpddtErrorDate.ToString <> "" Then sb.Append("dtErrorDate=" & sInsUpddtErrorDate & ",")
    If sInsUpdcErrorNumber.ToString <> "" Then sb.Append("cErrorNumber=" & sInsUpdcErrorNumber & ",")
    If sInsUpdcErrorSource.ToString <> "" Then sb.Append("cErrorSource=" & sInsUpdcErrorSource & ",")
    If sInsUpdcErrorMessage.ToString <> "" Then sb.Append("cErrorMessage=" & sInsUpdcErrorMessage & ",")
    If sInsUpdcErrorTargetSite.ToString <> "" Then sb.Append("cErrorTargetSite=" & sInsUpdcErrorTargetSite & ",")
    If sInsUpdcErrorStackTrace.ToString <> "" Then sb.Append("cErrorStackTrace=" & sInsUpdcErrorStackTrace & ",")
    If sInsUpdcErrorHTML.ToString <> "" Then sb.Append("cErrorHTML=" & sInsUpdcErrorHTML & ",")
    If sInsUpdcErrorUserInput.ToString <> "" Then sb.Append("cErrorUserInput=" & sInsUpdcErrorUserInput & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where idError=" & sInsUpdidError
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [WebErrorLog] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("idError=" & sInsUpdidError)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TableHostPayments

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iHostPaymentID As Int32
  Private sInsUpdHostPaymentID As String
  Property HostPaymentID_PK__Integer() As Int32
    Get
      Return iHostPaymentID
    End Get
    Set(ByVal Value As Int32)
      iHostPaymentID = Value
      sInsUpdHostPaymentID = oUtil.FixParam(iHostPaymentID, True)
    End Set
  End Property

  Private iHostID As Int32
  Private sInsUpdHostID As String
  Property HostID_RQ__Integer() As Int32
    Get
      Return iHostID
    End Get
    Set(ByVal Value As Int32)
      iHostID = Value
      sInsUpdHostID = oUtil.FixParam(iHostID, False)
    End Set
  End Property

  Private iFromBookingID As Int32
  Private sInsUpdFromBookingID As String
  Property FromBookingID__Integer() As Int32
    Get
      Return iFromBookingID
    End Get
    Set(ByVal Value As Int32)
      iFromBookingID = Value
      sInsUpdFromBookingID = oUtil.FixParam(iFromBookingID, True)
    End Set
  End Property

  Private iToBookingID As Int32
  Private sInsUpdToBookingID As String
  Property ToBookingID__Integer() As Int32
    Get
      Return iToBookingID
    End Get
    Set(ByVal Value As Int32)
      iToBookingID = Value
      sInsUpdToBookingID = oUtil.FixParam(iToBookingID, True)
    End Set
  End Property

  Private sTransactionType As String
  Private sInsUpdTransactionType As String
  Property TransactionType_RQ__String() As String
    Get
      Return sTransactionType
    End Get
    Set(ByVal Value As String)
      sTransactionType = Value
      sInsUpdTransactionType = oUtil.FixParam(sTransactionType, False)
    End Set
  End Property

  Private sTransactionDate As String
  Private sInsUpdTransactionDate As String
  Property TransactionDate_RQ__Date() As String
    Get
      Return sTransactionDate
    End Get
    Set(ByVal Value As String)
      sTransactionDate = Value
      sInsUpdTransactionDate = oUtil.FixParam(sTransactionDate, False)
    End Set
  End Property

  Private sUsername As String
  Private sInsUpdUsername As String
  Property Username__String() As String
    Get
      Return sUsername
    End Get
    Set(ByVal Value As String)
      sUsername = Value
      sInsUpdUsername = oUtil.FixParam(sUsername, True)
    End Set
  End Property

  Private dAmount As Double
  Private sInsUpdAmount As String
  Property Amount_RQ__Numeric() As Double
    Get
      Return dAmount
    End Get
    Set(ByVal Value As Double)
      dAmount = Value
      sInsUpdAmount = oUtil.FixParam(dAmount, False)
    End Set
  End Property

  Private sNotes As String
  Private sInsUpdNotes As String
  Property Notes__String() As String
    Get
      Return sNotes
    End Get
    Set(ByVal Value As String)
      sNotes = Value
      sInsUpdNotes = oUtil.FixParam(sNotes, True)
    End Set
  End Property

  Private iParentHostPaymentID As Int32
  Private sInsUpdParentHostPaymentID As String
  Property ParentHostPaymentID__Integer() As Int32
    Get
      Return iParentHostPaymentID
    End Get
    Set(ByVal Value As Int32)
      iParentHostPaymentID = Value
      sInsUpdParentHostPaymentID = oUtil.FixParam(iParentHostPaymentID, True)
    End Set
  End Property

  Private iParentHostPaymentBookingYear As Int32
  Private sInsUpdParentHostPaymentBookingYear As String
  Property ParentHostPaymentBookingYear__Integer() As Int32
    Get
      Return iParentHostPaymentBookingYear
    End Get
    Set(ByVal Value As Int32)
      iParentHostPaymentBookingYear = Value
      sInsUpdParentHostPaymentBookingYear = oUtil.FixParam(iParentHostPaymentBookingYear, True)
    End Set
  End Property

  Private iPropertyGroup As Int32
  Private sInsUpdPropertyGroup As String
  Property PropertyGroup__Integer() As Int32
    Get
      Return iPropertyGroup
    End Get
    Set(ByVal Value As Int32)
      iPropertyGroup = Value
      sInsUpdPropertyGroup = oUtil.FixParam(iPropertyGroup, True)
    End Set
  End Property

  Public Sub Clear()
    iHostPaymentID = 0
    sInsUpdHostPaymentID = ""
    iHostID = 0
    sInsUpdHostID = ""
    iFromBookingID = 0
    sInsUpdFromBookingID = ""
    iToBookingID = 0
    sInsUpdToBookingID = ""
    sTransactionType = ""
    sInsUpdTransactionType = ""
    sTransactionDate = ""
    sInsUpdTransactionDate = ""
    sUsername = ""
    sInsUpdUsername = ""
    dAmount = 0.0
    sInsUpdAmount = ""
    sNotes = ""
    sInsUpdNotes = ""
    iParentHostPaymentID = 0
    sInsUpdParentHostPaymentID = ""
    iParentHostPaymentBookingYear = 0
    sInsUpdParentHostPaymentBookingYear = ""
    iPropertyGroup = 0
    sInsUpdPropertyGroup = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdHostPaymentID.ToString = "'-12345'" Then sb.Append("HostPaymentID,")
        If sInsUpdHostID.ToString = "'-12345'" Then sb.Append("HostID,")
        If sInsUpdFromBookingID.ToString = "'-12345'" Then sb.Append("FromBookingID,")
        If sInsUpdToBookingID.ToString = "'-12345'" Then sb.Append("ToBookingID,")
        If sInsUpdTransactionType.ToString = "'Select'" Then sb.Append("TransactionType,")
        If sInsUpdTransactionDate.ToString = "'Select'" Then sb.Append("TransactionDate,")
        If sInsUpdUsername.ToString = "'Select'" Then sb.Append("Username,")
        If sInsUpdAmount.ToString = "'-12345'" Then sb.Append("Amount,")
        If sInsUpdNotes.ToString = "'Select'" Then sb.Append("Notes,")
        If sInsUpdParentHostPaymentID.ToString = "'-12345'" Then sb.Append("ParentHostPaymentID,")
        If sInsUpdParentHostPaymentBookingYear.ToString = "'-12345'" Then sb.Append("ParentHostPaymentBookingYear,")
        If sInsUpdPropertyGroup.ToString = "'-12345'" Then sb.Append("PropertyGroup,")
      Else
        sb.Append("HostPaymentID,")
        sb.Append("HostID,")
        sb.Append("FromBookingID,")
        sb.Append("ToBookingID,")
        sb.Append("TransactionType,")
        sb.Append("TransactionDate,")
        sb.Append("Username,")
        sb.Append("Amount,")
        sb.Append("Notes,")
        sb.Append("ParentHostPaymentID,")
        sb.Append("ParentHostPaymentBookingYear,")
        sb.Append("PropertyGroup,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [HostPayments]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdHostPaymentID.ToString <> "") And (sInsUpdHostPaymentID <> "'-12345'") Then sbw.Append("HostPaymentID=" & sInsUpdHostPaymentID & " and ")
      If (sInsUpdHostID.ToString <> "") And (sInsUpdHostID <> "'-12345'") Then sbw.Append("HostID=" & sInsUpdHostID & " and ")
      If (sInsUpdFromBookingID.ToString <> "") And (sInsUpdFromBookingID <> "'-12345'") Then sbw.Append("FromBookingID=" & sInsUpdFromBookingID & " and ")
      If (sInsUpdToBookingID.ToString <> "") And (sInsUpdToBookingID <> "'-12345'") Then sbw.Append("ToBookingID=" & sInsUpdToBookingID & " and ")
      If (sInsUpdTransactionType.ToString <> "") And (sInsUpdTransactionType <> "'Select'") Then sbw.Append("TransactionType=" & sInsUpdTransactionType & " and ")
      If (sInsUpdTransactionDate.ToString <> "") And (sInsUpdTransactionDate <> "'Select'") Then sbw.Append("TransactionDate=" & sInsUpdTransactionDate & " and ")
      If (sInsUpdUsername.ToString <> "") And (sInsUpdUsername <> "'Select'") Then sbw.Append("Username=" & sInsUpdUsername & " and ")
      If (sInsUpdAmount.ToString <> "") And (sInsUpdAmount <> "'-12345'") Then sbw.Append("Amount=" & sInsUpdAmount & " and ")
      If (sInsUpdNotes.ToString <> "") And (sInsUpdNotes <> "'Select'") Then sbw.Append("Notes=" & sInsUpdNotes & " and ")
      If (sInsUpdParentHostPaymentID.ToString <> "") And (sInsUpdParentHostPaymentID <> "'-12345'") Then sbw.Append("ParentHostPaymentID=" & sInsUpdParentHostPaymentID & " and ")
      If (sInsUpdParentHostPaymentBookingYear.ToString <> "") And (sInsUpdParentHostPaymentBookingYear <> "'-12345'") Then sbw.Append("ParentHostPaymentBookingYear=" & sInsUpdParentHostPaymentBookingYear & " and ")
      If (sInsUpdPropertyGroup.ToString <> "") And (sInsUpdPropertyGroup <> "'-12345'") Then sbw.Append("PropertyGroup=" & sInsUpdPropertyGroup & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      HostPaymentID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("HostPaymentID")), 0, CurrentRow.Item("HostPaymentID").ToString)
      HostID_RQ__Integer = IIf(IsDBNull(CurrentRow.Item("HostID")), 0, CurrentRow.Item("HostID"))
      FromBookingID__Integer = IIf(IsDBNull(CurrentRow.Item("FromBookingID")), 0, CurrentRow.Item("FromBookingID"))
      ToBookingID__Integer = IIf(IsDBNull(CurrentRow.Item("ToBookingID")), 0, CurrentRow.Item("ToBookingID"))
      TransactionType_RQ__String = IIf(IsDBNull(CurrentRow.Item("TransactionType")), "", CurrentRow.Item("TransactionType"))
      TransactionDate_RQ__Date = IIf(IsDBNull(CurrentRow.Item("TransactionDate")), "", CurrentRow.Item("TransactionDate"))
      Username__String = IIf(IsDBNull(CurrentRow.Item("Username")), "", CurrentRow.Item("Username"))
      Amount_RQ__Numeric = IIf(IsDBNull(CurrentRow.Item("Amount")), 0.0, CurrentRow.Item("Amount"))
      Notes__String = IIf(IsDBNull(CurrentRow.Item("Notes")), "", CurrentRow.Item("Notes"))
      ParentHostPaymentID__Integer = IIf(IsDBNull(CurrentRow.Item("ParentHostPaymentID")), 0, CurrentRow.Item("ParentHostPaymentID"))
      ParentHostPaymentBookingYear__Integer = IIf(IsDBNull(CurrentRow.Item("ParentHostPaymentBookingYear")), 0, CurrentRow.Item("ParentHostPaymentBookingYear"))
      PropertyGroup__Integer = IIf(IsDBNull(CurrentRow.Item("PropertyGroup")), 0, CurrentRow.Item("PropertyGroup"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [HostPayments](")
    If sInsUpdHostID.ToString <> "" Then
      sb.Append("HostID,")
      sbv.Append(sInsUpdHostID & ",")
    End If
    If sInsUpdFromBookingID.ToString <> "" Then
      sb.Append("FromBookingID,")
      sbv.Append(sInsUpdFromBookingID & ",")
    End If
    If sInsUpdToBookingID.ToString <> "" Then
      sb.Append("ToBookingID,")
      sbv.Append(sInsUpdToBookingID & ",")
    End If
    If sInsUpdTransactionType.ToString <> "" Then
      sb.Append("TransactionType,")
      sbv.Append(sInsUpdTransactionType & ",")
    End If
    If sInsUpdTransactionDate.ToString <> "" Then
      sb.Append("TransactionDate,")
      sbv.Append(sInsUpdTransactionDate & ",")
    End If
    If sInsUpdUsername.ToString <> "" Then
      sb.Append("Username,")
      sbv.Append(sInsUpdUsername & ",")
    End If
    If sInsUpdAmount.ToString <> "" Then
      sb.Append("Amount,")
      sbv.Append(sInsUpdAmount & ",")
    End If
    If sInsUpdNotes.ToString <> "" Then
      sb.Append("Notes,")
      sbv.Append(sInsUpdNotes & ",")
    End If
    If sInsUpdParentHostPaymentID.ToString <> "" Then
      sb.Append("ParentHostPaymentID,")
      sbv.Append(sInsUpdParentHostPaymentID & ",")
    End If
    If sInsUpdParentHostPaymentBookingYear.ToString <> "" Then
      sb.Append("ParentHostPaymentBookingYear,")
      sbv.Append(sInsUpdParentHostPaymentBookingYear & ",")
    End If
    If sInsUpdPropertyGroup.ToString <> "" Then
      sb.Append("PropertyGroup,")
      sbv.Append(sInsUpdPropertyGroup & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(HostPaymentID) from [HostPayments]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    HostPaymentID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [HostPayments] Set ")
    If sInsUpdHostID.ToString <> "" Then sb.Append("HostID=" & sInsUpdHostID & ",")
    If sInsUpdFromBookingID.ToString <> "" Then sb.Append("FromBookingID=" & sInsUpdFromBookingID & ",")
    If sInsUpdToBookingID.ToString <> "" Then sb.Append("ToBookingID=" & sInsUpdToBookingID & ",")
    If sInsUpdTransactionType.ToString <> "" Then sb.Append("TransactionType=" & sInsUpdTransactionType & ",")
    If sInsUpdTransactionDate.ToString <> "" Then sb.Append("TransactionDate=" & sInsUpdTransactionDate & ",")
    If sInsUpdUsername.ToString <> "" Then sb.Append("Username=" & sInsUpdUsername & ",")
    If sInsUpdAmount.ToString <> "" Then sb.Append("Amount=" & sInsUpdAmount & ",")
    If sInsUpdNotes.ToString <> "" Then sb.Append("Notes=" & sInsUpdNotes & ",")
    If sInsUpdParentHostPaymentID.ToString <> "" Then sb.Append("ParentHostPaymentID=" & sInsUpdParentHostPaymentID & ",")
    If sInsUpdParentHostPaymentBookingYear.ToString <> "" Then sb.Append("ParentHostPaymentBookingYear=" & sInsUpdParentHostPaymentBookingYear & ",")
    If sInsUpdPropertyGroup.ToString <> "" Then sb.Append("PropertyGroup=" & sInsUpdPropertyGroup & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where HostPaymentID=" & sInsUpdHostPaymentID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [HostPayments] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("HostPaymentID=" & sInsUpdHostPaymentID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TablePropertyCategories

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iCategory_ID As Int32
  Private sInsUpdCategory_ID As String
  Property Category_ID_PK__Integer() As Int32
    Get
      Return iCategory_ID
    End Get
    Set(ByVal Value As Int32)
      iCategory_ID = Value
      sInsUpdCategory_ID = oUtil.FixParam(iCategory_ID, True)
    End Set
  End Property

  Private sCategory As String
  Private sInsUpdCategory As String
  Property Category__String() As String
    Get
      Return sCategory
    End Get
    Set(ByVal Value As String)
      sCategory = Value
      sInsUpdCategory = oUtil.FixParam(sCategory, True)
    End Set
  End Property

  Private sWebCategoryTitle As String
  Private sInsUpdWebCategoryTitle As String
  Property WebCategoryTitle__String() As String
    Get
      Return sWebCategoryTitle
    End Get
    Set(ByVal Value As String)
      sWebCategoryTitle = Value
      sInsUpdWebCategoryTitle = oUtil.FixParam(sWebCategoryTitle, True)
    End Set
  End Property

  Private sWebLeftSection As String
  Private sInsUpdWebLeftSection As String
  Property WebLeftSection__String() As String
    Get
      Return sWebLeftSection
    End Get
    Set(ByVal Value As String)
      sWebLeftSection = Value
      sInsUpdWebLeftSection = oUtil.FixParam(sWebLeftSection, True)
    End Set
  End Property

  Private sWebMidSection As String
  Private sInsUpdWebMidSection As String
  Property WebMidSection__String() As String
    Get
      Return sWebMidSection
    End Get
    Set(ByVal Value As String)
      sWebMidSection = Value
      sInsUpdWebMidSection = oUtil.FixParam(sWebMidSection, True)
    End Set
  End Property

  Private sWebRightSection As String
  Private sInsUpdWebRightSection As String
  Property WebRightSection__String() As String
    Get
      Return sWebRightSection
    End Get
    Set(ByVal Value As String)
      sWebRightSection = Value
      sInsUpdWebRightSection = oUtil.FixParam(sWebRightSection, True)
    End Set
  End Property

  Public Sub Clear()
    iCategory_ID = 0
    sInsUpdCategory_ID = ""
    sCategory = ""
    sInsUpdCategory = ""
    sWebCategoryTitle = ""
    sInsUpdWebCategoryTitle = ""
    sWebLeftSection = ""
    sInsUpdWebLeftSection = ""
    sWebMidSection = ""
    sInsUpdWebMidSection = ""
    sWebRightSection = ""
    sInsUpdWebRightSection = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdCategory_ID.ToString = "'-12345'" Then sb.Append("Category_ID,")
        If sInsUpdCategory.ToString = "'Select'" Then sb.Append("Category,")
        If sInsUpdWebCategoryTitle.ToString = "'Select'" Then sb.Append("WebCategoryTitle,")
        If sInsUpdWebLeftSection.ToString = "'Select'" Then sb.Append("WebLeftSection,")
        If sInsUpdWebMidSection.ToString = "'Select'" Then sb.Append("WebMidSection,")
        If sInsUpdWebRightSection.ToString = "'Select'" Then sb.Append("WebRightSection,")
      Else
        sb.Append("Category_ID,")
        sb.Append("Category,")
        sb.Append("WebCategoryTitle,")
        sb.Append("WebLeftSection,")
        sb.Append("WebMidSection,")
        sb.Append("WebRightSection,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [PropertyCategories]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdCategory_ID.ToString <> "") And (sInsUpdCategory_ID <> "'-12345'") Then sbw.Append("Category_ID=" & sInsUpdCategory_ID & " and ")
      If (sInsUpdCategory.ToString <> "") And (sInsUpdCategory <> "'Select'") Then sbw.Append("Category=" & sInsUpdCategory & " and ")
      If (sInsUpdWebCategoryTitle.ToString <> "") And (sInsUpdWebCategoryTitle <> "'Select'") Then sbw.Append("WebCategoryTitle=" & sInsUpdWebCategoryTitle & " and ")
      If (sInsUpdWebLeftSection.ToString <> "") And (sInsUpdWebLeftSection <> "'Select'") Then sbw.Append("WebLeftSection=" & sInsUpdWebLeftSection & " and ")
      If (sInsUpdWebMidSection.ToString <> "") And (sInsUpdWebMidSection <> "'Select'") Then sbw.Append("WebMidSection=" & sInsUpdWebMidSection & " and ")
      If (sInsUpdWebRightSection.ToString <> "") And (sInsUpdWebRightSection <> "'Select'") Then sbw.Append("WebRightSection=" & sInsUpdWebRightSection & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      Category_ID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("Category_ID")), 0, CurrentRow.Item("Category_ID").ToString)
      Category__String = IIf(IsDBNull(CurrentRow.Item("Category")), "", CurrentRow.Item("Category"))
      WebCategoryTitle__String = IIf(IsDBNull(CurrentRow.Item("WebCategoryTitle")), "", CurrentRow.Item("WebCategoryTitle"))
      WebLeftSection__String = IIf(IsDBNull(CurrentRow.Item("WebLeftSection")), "", CurrentRow.Item("WebLeftSection"))
      WebMidSection__String = IIf(IsDBNull(CurrentRow.Item("WebMidSection")), "", CurrentRow.Item("WebMidSection"))
      WebRightSection__String = IIf(IsDBNull(CurrentRow.Item("WebRightSection")), "", CurrentRow.Item("WebRightSection"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [PropertyCategories](")
    If sInsUpdCategory.ToString <> "" Then
      sb.Append("Category,")
      sbv.Append(sInsUpdCategory & ",")
    End If
    If sInsUpdWebCategoryTitle.ToString <> "" Then
      sb.Append("WebCategoryTitle,")
      sbv.Append(sInsUpdWebCategoryTitle & ",")
    End If
    If sInsUpdWebLeftSection.ToString <> "" Then
      sb.Append("WebLeftSection,")
      sbv.Append(sInsUpdWebLeftSection & ",")
    End If
    If sInsUpdWebMidSection.ToString <> "" Then
      sb.Append("WebMidSection,")
      sbv.Append(sInsUpdWebMidSection & ",")
    End If
    If sInsUpdWebRightSection.ToString <> "" Then
      sb.Append("WebRightSection,")
      sbv.Append(sInsUpdWebRightSection & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(Category_ID) from [PropertyCategories]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    Category_ID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [PropertyCategories] Set ")
    If sInsUpdCategory.ToString <> "" Then sb.Append("Category=" & sInsUpdCategory & ",")
    If sInsUpdWebCategoryTitle.ToString <> "" Then sb.Append("WebCategoryTitle=" & sInsUpdWebCategoryTitle & ",")
    If sInsUpdWebLeftSection.ToString <> "" Then sb.Append("WebLeftSection=" & sInsUpdWebLeftSection & ",")
    If sInsUpdWebMidSection.ToString <> "" Then sb.Append("WebMidSection=" & sInsUpdWebMidSection & ",")
    If sInsUpdWebRightSection.ToString <> "" Then sb.Append("WebRightSection=" & sInsUpdWebRightSection & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where Category_ID=" & sInsUpdCategory_ID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [PropertyCategories] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("Category_ID=" & sInsUpdCategory_ID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TableStatus

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iRow_ID As Int32
  Private sInsUpdRow_ID As String
  Property Row_ID_PK__Integer() As Int32
    Get
      Return iRow_ID
    End Get
    Set(ByVal Value As Int32)
      iRow_ID = Value
      sInsUpdRow_ID = oUtil.FixParam(iRow_ID, True)
    End Set
  End Property

  Private sStatus As String
  Private sInsUpdStatus As String
  Property Status__String() As String
    Get
      Return sStatus
    End Get
    Set(ByVal Value As String)
      sStatus = Value
      sInsUpdStatus = oUtil.FixParam(sStatus, True)
    End Set
  End Property

  Public Sub Clear()
    iRow_ID = 0
    sInsUpdRow_ID = ""
    sStatus = ""
    sInsUpdStatus = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdRow_ID.ToString = "'-12345'" Then sb.Append("Row_ID,")
        If sInsUpdStatus.ToString = "'Select'" Then sb.Append("Status,")
      Else
        sb.Append("Row_ID,")
        sb.Append("Status,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [Status]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdRow_ID.ToString <> "") And (sInsUpdRow_ID <> "'-12345'") Then sbw.Append("Row_ID=" & sInsUpdRow_ID & " and ")
      If (sInsUpdStatus.ToString <> "") And (sInsUpdStatus <> "'Select'") Then sbw.Append("Status=" & sInsUpdStatus & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      Row_ID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("Row_ID")), 0, CurrentRow.Item("Row_ID").ToString)
      Status__String = IIf(IsDBNull(CurrentRow.Item("Status")), "", CurrentRow.Item("Status"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [Status](")
    If sInsUpdStatus.ToString <> "" Then
      sb.Append("Status,")
      sbv.Append(sInsUpdStatus & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(Row_ID) from [Status]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    Row_ID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [Status] Set ")
    If sInsUpdStatus.ToString <> "" Then sb.Append("Status=" & sInsUpdStatus & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where Row_ID=" & sInsUpdRow_ID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [Status] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("Row_ID=" & sInsUpdRow_ID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TableCIMGatewayActivity

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iActivityID As Int32
  Private sInsUpdActivityID As String
  Property ActivityID_PK__Integer() As Int32
    Get
      Return iActivityID
    End Get
    Set(ByVal Value As Int32)
      iActivityID = Value
      sInsUpdActivityID = oUtil.FixParam(iActivityID, True)
    End Set
  End Property

  Private sGatewayRequest As String
  Private sInsUpdGatewayRequest As String
  Property GatewayRequest__String() As String
    Get
      Return sGatewayRequest
    End Get
    Set(ByVal Value As String)
      sGatewayRequest = Value
      sInsUpdGatewayRequest = oUtil.FixParam(sGatewayRequest, True)
    End Set
  End Property

  Private sGatewayResponse As String
  Private sInsUpdGatewayResponse As String
  Property GatewayResponse__String() As String
    Get
      Return sGatewayResponse
    End Get
    Set(ByVal Value As String)
      sGatewayResponse = Value
      sInsUpdGatewayResponse = oUtil.FixParam(sGatewayResponse, True)
    End Set
  End Property

  Private sGatewayReturn As String
  Private sInsUpdGatewayReturn As String
  Property GatewayReturn__String() As String
    Get
      Return sGatewayReturn
    End Get
    Set(ByVal Value As String)
      sGatewayReturn = Value
      sInsUpdGatewayReturn = oUtil.FixParam(sGatewayReturn, True)
    End Set
  End Property

  Private sGatewayError As String
  Private sInsUpdGatewayError As String
  Property GatewayError__String() As String
    Get
      Return sGatewayError
    End Get
    Set(ByVal Value As String)
      sGatewayError = Value
      sInsUpdGatewayError = oUtil.FixParam(sGatewayError, True)
    End Set
  End Property

  Private iGuestID As Int32
  Private sInsUpdGuestID As String
  Property GuestID__Integer() As Int32
    Get
      Return iGuestID
    End Get
    Set(ByVal Value As Int32)
      iGuestID = Value
      sInsUpdGuestID = oUtil.FixParam(iGuestID, True)
    End Set
  End Property

  Private iBookingID As Int32
  Private sInsUpdBookingID As String
  Property BookingID__Integer() As Int32
    Get
      Return iBookingID
    End Get
    Set(ByVal Value As Int32)
      iBookingID = Value
      sInsUpdBookingID = oUtil.FixParam(iBookingID, True)
    End Set
  End Property

  Private iPaymentID As Int32
  Private sInsUpdPaymentID As String
  Property PaymentID__Integer() As Int32
    Get
      Return iPaymentID
    End Get
    Set(ByVal Value As Int32)
      iPaymentID = Value
      sInsUpdPaymentID = oUtil.FixParam(iPaymentID, True)
    End Set
  End Property

  Private sPaymentCategory As String
  Private sInsUpdPaymentCategory As String
  Property PaymentCategory__String() As String
    Get
      Return sPaymentCategory
    End Get
    Set(ByVal Value As String)
      sPaymentCategory = Value
      sInsUpdPaymentCategory = oUtil.FixParam(sPaymentCategory, True)
    End Set
  End Property

  Public Sub Clear()
    iActivityID = 0
    sInsUpdActivityID = ""
    sGatewayRequest = ""
    sInsUpdGatewayRequest = ""
    sGatewayResponse = ""
    sInsUpdGatewayResponse = ""
    sGatewayReturn = ""
    sInsUpdGatewayReturn = ""
    sGatewayError = ""
    sInsUpdGatewayError = ""
    iGuestID = 0
    sInsUpdGuestID = ""
    iBookingID = 0
    sInsUpdBookingID = ""
    iPaymentID = 0
    sInsUpdPaymentID = ""
    sPaymentCategory = ""
    sInsUpdPaymentCategory = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdActivityID.ToString = "'-12345'" Then sb.Append("ActivityID,")
        If sInsUpdGatewayRequest.ToString = "'Select'" Then sb.Append("GatewayRequest,")
        If sInsUpdGatewayResponse.ToString = "'Select'" Then sb.Append("GatewayResponse,")
        If sInsUpdGatewayReturn.ToString = "'Select'" Then sb.Append("GatewayReturn,")
        If sInsUpdGatewayError.ToString = "'Select'" Then sb.Append("GatewayError,")
        If sInsUpdGuestID.ToString = "'-12345'" Then sb.Append("GuestID,")
        If sInsUpdBookingID.ToString = "'-12345'" Then sb.Append("BookingID,")
        If sInsUpdPaymentID.ToString = "'-12345'" Then sb.Append("PaymentID,")
        If sInsUpdPaymentCategory.ToString = "'Select'" Then sb.Append("PaymentCategory,")
      Else
        sb.Append("ActivityID,")
        sb.Append("GatewayRequest,")
        sb.Append("GatewayResponse,")
        sb.Append("GatewayReturn,")
        sb.Append("GatewayError,")
        sb.Append("GuestID,")
        sb.Append("BookingID,")
        sb.Append("PaymentID,")
        sb.Append("PaymentCategory,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [CIMGatewayActivity]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdActivityID.ToString <> "") And (sInsUpdActivityID <> "'-12345'") Then sbw.Append("ActivityID=" & sInsUpdActivityID & " and ")
      If (sInsUpdGatewayRequest.ToString <> "") And (sInsUpdGatewayRequest <> "'Select'") Then sbw.Append("GatewayRequest=" & sInsUpdGatewayRequest & " and ")
      If (sInsUpdGatewayResponse.ToString <> "") And (sInsUpdGatewayResponse <> "'Select'") Then sbw.Append("GatewayResponse=" & sInsUpdGatewayResponse & " and ")
      If (sInsUpdGatewayReturn.ToString <> "") And (sInsUpdGatewayReturn <> "'Select'") Then sbw.Append("GatewayReturn=" & sInsUpdGatewayReturn & " and ")
      If (sInsUpdGatewayError.ToString <> "") And (sInsUpdGatewayError <> "'Select'") Then sbw.Append("GatewayError=" & sInsUpdGatewayError & " and ")
      If (sInsUpdGuestID.ToString <> "") And (sInsUpdGuestID <> "'-12345'") Then sbw.Append("GuestID=" & sInsUpdGuestID & " and ")
      If (sInsUpdBookingID.ToString <> "") And (sInsUpdBookingID <> "'-12345'") Then sbw.Append("BookingID=" & sInsUpdBookingID & " and ")
      If (sInsUpdPaymentID.ToString <> "") And (sInsUpdPaymentID <> "'-12345'") Then sbw.Append("PaymentID=" & sInsUpdPaymentID & " and ")
      If (sInsUpdPaymentCategory.ToString <> "") And (sInsUpdPaymentCategory <> "'Select'") Then sbw.Append("PaymentCategory=" & sInsUpdPaymentCategory & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      ActivityID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("ActivityID")), 0, CurrentRow.Item("ActivityID").ToString)
      GatewayRequest__String = IIf(IsDBNull(CurrentRow.Item("GatewayRequest")), "", CurrentRow.Item("GatewayRequest"))
      GatewayResponse__String = IIf(IsDBNull(CurrentRow.Item("GatewayResponse")), "", CurrentRow.Item("GatewayResponse"))
      GatewayReturn__String = IIf(IsDBNull(CurrentRow.Item("GatewayReturn")), "", CurrentRow.Item("GatewayReturn"))
      GatewayError__String = IIf(IsDBNull(CurrentRow.Item("GatewayError")), "", CurrentRow.Item("GatewayError"))
      GuestID__Integer = IIf(IsDBNull(CurrentRow.Item("GuestID")), 0, CurrentRow.Item("GuestID"))
      BookingID__Integer = IIf(IsDBNull(CurrentRow.Item("BookingID")), 0, CurrentRow.Item("BookingID"))
      PaymentID__Integer = IIf(IsDBNull(CurrentRow.Item("PaymentID")), 0, CurrentRow.Item("PaymentID"))
      PaymentCategory__String = IIf(IsDBNull(CurrentRow.Item("PaymentCategory")), "", CurrentRow.Item("PaymentCategory"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [CIMGatewayActivity](")
    If sInsUpdGatewayRequest.ToString <> "" Then
      sb.Append("GatewayRequest,")
      sbv.Append(sInsUpdGatewayRequest & ",")
    End If
    If sInsUpdGatewayResponse.ToString <> "" Then
      sb.Append("GatewayResponse,")
      sbv.Append(sInsUpdGatewayResponse & ",")
    End If
    If sInsUpdGatewayReturn.ToString <> "" Then
      sb.Append("GatewayReturn,")
      sbv.Append(sInsUpdGatewayReturn & ",")
    End If
    If sInsUpdGatewayError.ToString <> "" Then
      sb.Append("GatewayError,")
      sbv.Append(sInsUpdGatewayError & ",")
    End If
    If sInsUpdGuestID.ToString <> "" Then
      sb.Append("GuestID,")
      sbv.Append(sInsUpdGuestID & ",")
    End If
    If sInsUpdBookingID.ToString <> "" Then
      sb.Append("BookingID,")
      sbv.Append(sInsUpdBookingID & ",")
    End If
    If sInsUpdPaymentID.ToString <> "" Then
      sb.Append("PaymentID,")
      sbv.Append(sInsUpdPaymentID & ",")
    End If
    If sInsUpdPaymentCategory.ToString <> "" Then
      sb.Append("PaymentCategory,")
      sbv.Append(sInsUpdPaymentCategory & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(ActivityID) from [CIMGatewayActivity]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    ActivityID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [CIMGatewayActivity] Set ")
    If sInsUpdGatewayRequest.ToString <> "" Then sb.Append("GatewayRequest=" & sInsUpdGatewayRequest & ",")
    If sInsUpdGatewayResponse.ToString <> "" Then sb.Append("GatewayResponse=" & sInsUpdGatewayResponse & ",")
    If sInsUpdGatewayReturn.ToString <> "" Then sb.Append("GatewayReturn=" & sInsUpdGatewayReturn & ",")
    If sInsUpdGatewayError.ToString <> "" Then sb.Append("GatewayError=" & sInsUpdGatewayError & ",")
    If sInsUpdGuestID.ToString <> "" Then sb.Append("GuestID=" & sInsUpdGuestID & ",")
    If sInsUpdBookingID.ToString <> "" Then sb.Append("BookingID=" & sInsUpdBookingID & ",")
    If sInsUpdPaymentID.ToString <> "" Then sb.Append("PaymentID=" & sInsUpdPaymentID & ",")
    If sInsUpdPaymentCategory.ToString <> "" Then sb.Append("PaymentCategory=" & sInsUpdPaymentCategory & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where ActivityID=" & sInsUpdActivityID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [CIMGatewayActivity] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("ActivityID=" & sInsUpdActivityID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TableBalancePayments

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iPaymentID As Int32
  Private sInsUpdPaymentID As String
  Property PaymentID_PK__Integer() As Int32
    Get
      Return iPaymentID
    End Get
    Set(ByVal Value As Int32)
      iPaymentID = Value
      sInsUpdPaymentID = oUtil.FixParam(iPaymentID, True)
    End Set
  End Property

  Private iBookingID As Int32
  Private sInsUpdBookingID As String
  Property BookingID__Integer() As Int32
    Get
      Return iBookingID
    End Get
    Set(ByVal Value As Int32)
      iBookingID = Value
      sInsUpdBookingID = oUtil.FixParam(iBookingID, True)
    End Set
  End Property

  Private iPropertyID As Int32
  Private sInsUpdPropertyID As String
  Property PropertyID__Integer() As Int32
    Get
      Return iPropertyID
    End Get
    Set(ByVal Value As Int32)
      iPropertyID = Value
      sInsUpdPropertyID = oUtil.FixParam(iPropertyID, True)
    End Set
  End Property

  Private iHostID As Int32
  Private sInsUpdHostID As String
  Property HostID__Integer() As Int32
    Get
      Return iHostID
    End Get
    Set(ByVal Value As Int32)
      iHostID = Value
      sInsUpdHostID = oUtil.FixParam(iHostID, True)
    End Set
  End Property

  Private iGuestID As Int32
  Private sInsUpdGuestID As String
  Property GuestID__Integer() As Int32
    Get
      Return iGuestID
    End Get
    Set(ByVal Value As Int32)
      iGuestID = Value
      sInsUpdGuestID = oUtil.FixParam(iGuestID, True)
    End Set
  End Property

  Private sCategory As String
  Private sInsUpdCategory As String
  Property Category__String() As String
    Get
      Return sCategory
    End Get
    Set(ByVal Value As String)
      sCategory = Value
      sInsUpdCategory = oUtil.FixParam(sCategory, True)
    End Set
  End Property

  Private sType As String
  Private sInsUpdType As String
  Property Type__String() As String
    Get
      Return sType
    End Get
    Set(ByVal Value As String)
      sType = Value
      sInsUpdType = oUtil.FixParam(sType, True)
    End Set
  End Property

  Private dAmount As Double
  Private sInsUpdAmount As String
  Property Amount__Numeric() As Double
    Get
      Return dAmount
    End Get
    Set(ByVal Value As Double)
      dAmount = Value
      sInsUpdAmount = oUtil.FixParam(dAmount, True)
    End Set
  End Property

  Private sDate As String
  Private sInsUpdDate As String
  Property Date__Date() As String
    Get
      Return sDate
    End Get
    Set(ByVal Value As String)
      sDate = Value
      sInsUpdDate = oUtil.FixParam(sDate, True)
    End Set
  End Property

  Private sComputer As String
  Private sInsUpdComputer As String
  Property Computer__String() As String
    Get
      Return sComputer
    End Get
    Set(ByVal Value As String)
      sComputer = Value
      sInsUpdComputer = oUtil.FixParam(sComputer, True)
    End Set
  End Property

  Private sCheckNumber As String
  Private sInsUpdCheckNumber As String
  Property CheckNumber__String() As String
    Get
      Return sCheckNumber
    End Get
    Set(ByVal Value As String)
      sCheckNumber = Value
      sInsUpdCheckNumber = oUtil.FixParam(sCheckNumber, True)
    End Set
  End Property

  Private sCheckName As String
  Private sInsUpdCheckName As String
  Property CheckName__String() As String
    Get
      Return sCheckName
    End Get
    Set(ByVal Value As String)
      sCheckName = Value
      sInsUpdCheckName = oUtil.FixParam(sCheckName, True)
    End Set
  End Property

  Private sCCType As String
  Private sInsUpdCCType As String
  Property CCType__String() As String
    Get
      Return sCCType
    End Get
    Set(ByVal Value As String)
      sCCType = Value
      sInsUpdCCType = oUtil.FixParam(sCCType, True)
    End Set
  End Property

  Private sCCNumber As String
  Private sInsUpdCCNumber As String
  Property CCNumber__String() As String
    Get
      Return sCCNumber
    End Get
    Set(ByVal Value As String)
      sCCNumber = Value
      sInsUpdCCNumber = oUtil.FixParam(sCCNumber, True)
    End Set
  End Property

  Private sCCNumberEncrypted As String
  Private sInsUpdCCNumberEncrypted As String
  Property CCNumberEncrypted__String() As String
    Get
      Return sCCNumberEncrypted
    End Get
    Set(ByVal Value As String)
      sCCNumberEncrypted = Value
      sInsUpdCCNumberEncrypted = oUtil.FixParam(sCCNumberEncrypted, True)
    End Set
  End Property

  Private sCCName As String
  Private sInsUpdCCName As String
  Property CCName__String() As String
    Get
      Return sCCName
    End Get
    Set(ByVal Value As String)
      sCCName = Value
      sInsUpdCCName = oUtil.FixParam(sCCName, True)
    End Set
  End Property

  Private sCCExpMonth As String
  Private sInsUpdCCExpMonth As String
  Property CCExpMonth__String() As String
    Get
      Return sCCExpMonth
    End Get
    Set(ByVal Value As String)
      sCCExpMonth = Value
      sInsUpdCCExpMonth = oUtil.FixParam(sCCExpMonth, True)
    End Set
  End Property

  Private sCCExpYear As String
  Private sInsUpdCCExpYear As String
  Property CCExpYear__String() As String
    Get
      Return sCCExpYear
    End Get
    Set(ByVal Value As String)
      sCCExpYear = Value
      sInsUpdCCExpYear = oUtil.FixParam(sCCExpYear, True)
    End Set
  End Property

  Private sCCVerification As String
  Private sInsUpdCCVerification As String
  Property CCVerification__String() As String
    Get
      Return sCCVerification
    End Get
    Set(ByVal Value As String)
      sCCVerification = Value
      sInsUpdCCVerification = oUtil.FixParam(sCCVerification, True)
    End Set
  End Property

  Private sCCConfirm As String
  Private sInsUpdCCConfirm As String
  Property CCConfirm__String() As String
    Get
      Return sCCConfirm
    End Get
    Set(ByVal Value As String)
      sCCConfirm = Value
      sInsUpdCCConfirm = oUtil.FixParam(sCCConfirm, True)
    End Set
  End Property

  Private sQBStatus As String
  Private sInsUpdQBStatus As String
  Property QBStatus__String() As String
    Get
      Return sQBStatus
    End Get
    Set(ByVal Value As String)
      sQBStatus = Value
      sInsUpdQBStatus = oUtil.FixParam(sQBStatus, True)
    End Set
  End Property

  Private sNotes As String
  Private sInsUpdNotes As String
  Property Notes__String() As String
    Get
      Return sNotes
    End Get
    Set(ByVal Value As String)
      sNotes = Value
      sInsUpdNotes = oUtil.FixParam(sNotes, True)
    End Set
  End Property

  Private sUser As String
  Private sInsUpdUser As String
  Property User__String() As String
    Get
      Return sUser
    End Get
    Set(ByVal Value As String)
      sUser = Value
      sInsUpdUser = oUtil.FixParam(sUser, True)
    End Set
  End Property

  Private sStatus As String
  Private sInsUpdStatus As String
  Property Status__String() As String
    Get
      Return sStatus
    End Get
    Set(ByVal Value As String)
      sStatus = Value
      sInsUpdStatus = oUtil.FixParam(sStatus, True)
    End Set
  End Property

  Private iUseOnReceipt As Int32
  Private sInsUpdUseOnReceipt As String
  Property UseOnReceipt__Integer() As Int32
    Get
      Return iUseOnReceipt
    End Get
    Set(ByVal Value As Int32)
      iUseOnReceipt = Value
      sInsUpdUseOnReceipt = oUtil.FixParam(iUseOnReceipt, True)
    End Set
  End Property

  Private sOrigCCNumber As String
  Private sInsUpdOrigCCNumber As String
  Property OrigCCNumber__String() As String
    Get
      Return sOrigCCNumber
    End Get
    Set(ByVal Value As String)
      sOrigCCNumber = Value
      sInsUpdOrigCCNumber = oUtil.FixParam(sOrigCCNumber, True)
    End Set
  End Property

  Private sCCAddress As String
  Private sInsUpdCCAddress As String
  Property CCAddress__String() As String
    Get
      Return sCCAddress
    End Get
    Set(ByVal Value As String)
      sCCAddress = Value
      sInsUpdCCAddress = oUtil.FixParam(sCCAddress, True)
    End Set
  End Property

  Private sCCCity As String
  Private sInsUpdCCCity As String
  Property CCCity__String() As String
    Get
      Return sCCCity
    End Get
    Set(ByVal Value As String)
      sCCCity = Value
      sInsUpdCCCity = oUtil.FixParam(sCCCity, True)
    End Set
  End Property

  Private sCCState As String
  Private sInsUpdCCState As String
  Property CCState__String() As String
    Get
      Return sCCState
    End Get
    Set(ByVal Value As String)
      sCCState = Value
      sInsUpdCCState = oUtil.FixParam(sCCState, True)
    End Set
  End Property

  Private sCCZip As String
  Private sInsUpdCCZip As String
  Property CCZip__String() As String
    Get
      Return sCCZip
    End Get
    Set(ByVal Value As String)
      sCCZip = Value
      sInsUpdCCZip = oUtil.FixParam(sCCZip, True)
    End Set
  End Property

  Private sAuthorizeNetReturnCode As String
  Private sInsUpdAuthorizeNetReturnCode As String
  Property AuthorizeNetReturnCode__String() As String
    Get
      Return sAuthorizeNetReturnCode
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetReturnCode = Value
      sInsUpdAuthorizeNetReturnCode = oUtil.FixParam(sAuthorizeNetReturnCode, True)
    End Set
  End Property

  Private sAuthorizeNetReasonCode As String
  Private sInsUpdAuthorizeNetReasonCode As String
  Property AuthorizeNetReasonCode__String() As String
    Get
      Return sAuthorizeNetReasonCode
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetReasonCode = Value
      sInsUpdAuthorizeNetReasonCode = oUtil.FixParam(sAuthorizeNetReasonCode, True)
    End Set
  End Property

  Private sAuthorizeNetReasonText As String
  Private sInsUpdAuthorizeNetReasonText As String
  Property AuthorizeNetReasonText__String() As String
    Get
      Return sAuthorizeNetReasonText
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetReasonText = Value
      sInsUpdAuthorizeNetReasonText = oUtil.FixParam(sAuthorizeNetReasonText, True)
    End Set
  End Property

  Private sAuthorizeNetApprovalCode As String
  Private sInsUpdAuthorizeNetApprovalCode As String
  Property AuthorizeNetApprovalCode__String() As String
    Get
      Return sAuthorizeNetApprovalCode
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetApprovalCode = Value
      sInsUpdAuthorizeNetApprovalCode = oUtil.FixParam(sAuthorizeNetApprovalCode, True)
    End Set
  End Property

  Private sAuthorizeNetAVSResultCode As String
  Private sInsUpdAuthorizeNetAVSResultCode As String
  Property AuthorizeNetAVSResultCode__String() As String
    Get
      Return sAuthorizeNetAVSResultCode
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetAVSResultCode = Value
      sInsUpdAuthorizeNetAVSResultCode = oUtil.FixParam(sAuthorizeNetAVSResultCode, True)
    End Set
  End Property

  Private sAuthorizeNetAVSResultText As String
  Private sInsUpdAuthorizeNetAVSResultText As String
  Property AuthorizeNetAVSResultText__String() As String
    Get
      Return sAuthorizeNetAVSResultText
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetAVSResultText = Value
      sInsUpdAuthorizeNetAVSResultText = oUtil.FixParam(sAuthorizeNetAVSResultText, True)
    End Set
  End Property

  Private sAuthorizeNetTransactionID As String
  Private sInsUpdAuthorizeNetTransactionID As String
  Property AuthorizeNetTransactionID__String() As String
    Get
      Return sAuthorizeNetTransactionID
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetTransactionID = Value
      sInsUpdAuthorizeNetTransactionID = oUtil.FixParam(sAuthorizeNetTransactionID, True)
    End Set
  End Property

  Private sAuthorizeNetCVCCResponseCode As String
  Private sInsUpdAuthorizeNetCVCCResponseCode As String
  Property AuthorizeNetCVCCResponseCode__String() As String
    Get
      Return sAuthorizeNetCVCCResponseCode
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetCVCCResponseCode = Value
      sInsUpdAuthorizeNetCVCCResponseCode = oUtil.FixParam(sAuthorizeNetCVCCResponseCode, True)
    End Set
  End Property

  Private sAuthorizeNetCVCCResponseText As String
  Private sInsUpdAuthorizeNetCVCCResponseText As String
  Property AuthorizeNetCVCCResponseText__String() As String
    Get
      Return sAuthorizeNetCVCCResponseText
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetCVCCResponseText = Value
      sInsUpdAuthorizeNetCVCCResponseText = oUtil.FixParam(sAuthorizeNetCVCCResponseText, True)
    End Set
  End Property

  Private sLocation As String
  Private sInsUpdLocation As String
  Property Location__String() As String
    Get
      Return sLocation
    End Get
    Set(ByVal Value As String)
      sLocation = Value
      sInsUpdLocation = oUtil.FixParam(sLocation, True)
    End Set
  End Property

  Private sBankName As String
  Private sInsUpdBankName As String
  Property BankName__String() As String
    Get
      Return sBankName
    End Get
    Set(ByVal Value As String)
      sBankName = Value
      sInsUpdBankName = oUtil.FixParam(sBankName, True)
    End Set
  End Property

  Private sBankAccountType As String
  Private sInsUpdBankAccountType As String
  Property BankAccountType__String() As String
    Get
      Return sBankAccountType
    End Get
    Set(ByVal Value As String)
      sBankAccountType = Value
      sInsUpdBankAccountType = oUtil.FixParam(sBankAccountType, True)
    End Set
  End Property

  Private sBankAccountName As String
  Private sInsUpdBankAccountName As String
  Property BankAccountName__String() As String
    Get
      Return sBankAccountName
    End Get
    Set(ByVal Value As String)
      sBankAccountName = Value
      sInsUpdBankAccountName = oUtil.FixParam(sBankAccountName, True)
    End Set
  End Property

  Private sBankAccountNumber As String
  Private sInsUpdBankAccountNumber As String
  Property BankAccountNumber__String() As String
    Get
      Return sBankAccountNumber
    End Get
    Set(ByVal Value As String)
      sBankAccountNumber = Value
      sInsUpdBankAccountNumber = oUtil.FixParam(sBankAccountNumber, True)
    End Set
  End Property

  Private sBankAccountNumberEncrypted As String
  Private sInsUpdBankAccountNumberEncrypted As String
  Property BankAccountNumberEncrypted__String() As String
    Get
      Return sBankAccountNumberEncrypted
    End Get
    Set(ByVal Value As String)
      sBankAccountNumberEncrypted = Value
      sInsUpdBankAccountNumberEncrypted = oUtil.FixParam(sBankAccountNumberEncrypted, True)
    End Set
  End Property

  Private sBankRoutingNumber As String
  Private sInsUpdBankRoutingNumber As String
  Property BankRoutingNumber__String() As String
    Get
      Return sBankRoutingNumber
    End Get
    Set(ByVal Value As String)
      sBankRoutingNumber = Value
      sInsUpdBankRoutingNumber = oUtil.FixParam(sBankRoutingNumber, True)
    End Set
  End Property

  Private sEmail As String
  Private sInsUpdEmail As String
  Property Email__String() As String
    Get
      Return sEmail
    End Get
    Set(ByVal Value As String)
      sEmail = Value
      sInsUpdEmail = oUtil.FixParam(sEmail, True)
    End Set
  End Property

  Private iQBActivityID As Int32
  Private sInsUpdQBActivityID As String
  Property QBActivityID__Integer() As Int32
    Get
      Return iQBActivityID
    End Get
    Set(ByVal Value As Int32)
      iQBActivityID = Value
      sInsUpdQBActivityID = oUtil.FixParam(iQBActivityID, True)
    End Set
  End Property

  Private sAuthorizeNetURL As String
  Private sInsUpdAuthorizeNetURL As String
  Property AuthorizeNetURL__String() As String
    Get
      Return sAuthorizeNetURL
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetURL = Value
      sInsUpdAuthorizeNetURL = oUtil.FixParam(sAuthorizeNetURL, True)
    End Set
  End Property

  Private sAuthNETPaymentProfileID As String
  Private sInsUpdAuthNETPaymentProfileID As String
  Property AuthNETPaymentProfileID__String() As String
    Get
      Return sAuthNETPaymentProfileID
    End Get
    Set(ByVal Value As String)
      sAuthNETPaymentProfileID = Value
      sInsUpdAuthNETPaymentProfileID = oUtil.FixParam(sAuthNETPaymentProfileID, True)
    End Set
  End Property

  Public Sub Clear()
    iPaymentID = 0
    sInsUpdPaymentID = ""
    iBookingID = 0
    sInsUpdBookingID = ""
    iPropertyID = 0
    sInsUpdPropertyID = ""
    iHostID = 0
    sInsUpdHostID = ""
    iGuestID = 0
    sInsUpdGuestID = ""
    sCategory = ""
    sInsUpdCategory = ""
    sType = ""
    sInsUpdType = ""
    dAmount = 0.0
    sInsUpdAmount = ""
    sDate = ""
    sInsUpdDate = ""
    sComputer = ""
    sInsUpdComputer = ""
    sCheckNumber = ""
    sInsUpdCheckNumber = ""
    sCheckName = ""
    sInsUpdCheckName = ""
    sCCType = ""
    sInsUpdCCType = ""
    sCCNumber = ""
    sInsUpdCCNumber = ""
    sCCNumberEncrypted = ""
    sInsUpdCCNumberEncrypted = ""
    sCCName = ""
    sInsUpdCCName = ""
    sCCExpMonth = ""
    sInsUpdCCExpMonth = ""
    sCCExpYear = ""
    sInsUpdCCExpYear = ""
    sCCVerification = ""
    sInsUpdCCVerification = ""
    sCCConfirm = ""
    sInsUpdCCConfirm = ""
    sQBStatus = ""
    sInsUpdQBStatus = ""
    sNotes = ""
    sInsUpdNotes = ""
    sUser = ""
    sInsUpdUser = ""
    sStatus = ""
    sInsUpdStatus = ""
    iUseOnReceipt = 0
    sInsUpdUseOnReceipt = ""
    sOrigCCNumber = ""
    sInsUpdOrigCCNumber = ""
    sCCAddress = ""
    sInsUpdCCAddress = ""
    sCCCity = ""
    sInsUpdCCCity = ""
    sCCState = ""
    sInsUpdCCState = ""
    sCCZip = ""
    sInsUpdCCZip = ""
    sAuthorizeNetReturnCode = ""
    sInsUpdAuthorizeNetReturnCode = ""
    sAuthorizeNetReasonCode = ""
    sInsUpdAuthorizeNetReasonCode = ""
    sAuthorizeNetReasonText = ""
    sInsUpdAuthorizeNetReasonText = ""
    sAuthorizeNetApprovalCode = ""
    sInsUpdAuthorizeNetApprovalCode = ""
    sAuthorizeNetAVSResultCode = ""
    sInsUpdAuthorizeNetAVSResultCode = ""
    sAuthorizeNetAVSResultText = ""
    sInsUpdAuthorizeNetAVSResultText = ""
    sAuthorizeNetTransactionID = ""
    sInsUpdAuthorizeNetTransactionID = ""
    sAuthorizeNetCVCCResponseCode = ""
    sInsUpdAuthorizeNetCVCCResponseCode = ""
    sAuthorizeNetCVCCResponseText = ""
    sInsUpdAuthorizeNetCVCCResponseText = ""
    sLocation = ""
    sInsUpdLocation = ""
    sBankName = ""
    sInsUpdBankName = ""
    sBankAccountType = ""
    sInsUpdBankAccountType = ""
    sBankAccountName = ""
    sInsUpdBankAccountName = ""
    sBankAccountNumber = ""
    sInsUpdBankAccountNumber = ""
    sBankAccountNumberEncrypted = ""
    sInsUpdBankAccountNumberEncrypted = ""
    sBankRoutingNumber = ""
    sInsUpdBankRoutingNumber = ""
    sEmail = ""
    sInsUpdEmail = ""
    iQBActivityID = 0
    sInsUpdQBActivityID = ""
    sAuthorizeNetURL = ""
    sInsUpdAuthorizeNetURL = ""
    sAuthNETPaymentProfileID = ""
    sInsUpdAuthNETPaymentProfileID = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdPaymentID.ToString = "'-12345'" Then sb.Append("PaymentID,")
        If sInsUpdBookingID.ToString = "'-12345'" Then sb.Append("BookingID,")
        If sInsUpdPropertyID.ToString = "'-12345'" Then sb.Append("PropertyID,")
        If sInsUpdHostID.ToString = "'-12345'" Then sb.Append("HostID,")
        If sInsUpdGuestID.ToString = "'-12345'" Then sb.Append("GuestID,")
        If sInsUpdCategory.ToString = "'Select'" Then sb.Append("Category,")
        If sInsUpdType.ToString = "'Select'" Then sb.Append("Type,")
        If sInsUpdAmount.ToString = "'-12345'" Then sb.Append("Amount,")
        If sInsUpdDate.ToString = "'Select'" Then sb.Append("[Date],")
        If sInsUpdComputer.ToString = "'Select'" Then sb.Append("Computer,")
        If sInsUpdCheckNumber.ToString = "'Select'" Then sb.Append("CheckNumber,")
        If sInsUpdCheckName.ToString = "'Select'" Then sb.Append("CheckName,")
        If sInsUpdCCType.ToString = "'Select'" Then sb.Append("CCType,")
        If sInsUpdCCNumber.ToString = "'Select'" Then sb.Append("CCNumber,")
        If sInsUpdCCNumberEncrypted.ToString = "'Select'" Then sb.Append("CCNumberEncrypted,")
        If sInsUpdCCName.ToString = "'Select'" Then sb.Append("CCName,")
        If sInsUpdCCExpMonth.ToString = "'Select'" Then sb.Append("CCExpMonth,")
        If sInsUpdCCExpYear.ToString = "'Select'" Then sb.Append("CCExpYear,")
        If sInsUpdCCVerification.ToString = "'Select'" Then sb.Append("CCVerification,")
        If sInsUpdCCConfirm.ToString = "'Select'" Then sb.Append("CCConfirm,")
        If sInsUpdQBStatus.ToString = "'Select'" Then sb.Append("QBStatus,")
        If sInsUpdNotes.ToString = "'Select'" Then sb.Append("Notes,")
        If sInsUpdUser.ToString = "'Select'" Then sb.Append("[User],")
        If sInsUpdStatus.ToString = "'Select'" Then sb.Append("Status,")
        If sInsUpdUseOnReceipt.ToString = "'-12345'" Then sb.Append("UseOnReceipt,")
        If sInsUpdOrigCCNumber.ToString = "'Select'" Then sb.Append("OrigCCNumber,")
        If sInsUpdCCAddress.ToString = "'Select'" Then sb.Append("CCAddress,")
        If sInsUpdCCCity.ToString = "'Select'" Then sb.Append("CCCity,")
        If sInsUpdCCState.ToString = "'Select'" Then sb.Append("CCState,")
        If sInsUpdCCZip.ToString = "'Select'" Then sb.Append("CCZip,")
        If sInsUpdAuthorizeNetReturnCode.ToString = "'Select'" Then sb.Append("AuthorizeNetReturnCode,")
        If sInsUpdAuthorizeNetReasonCode.ToString = "'Select'" Then sb.Append("AuthorizeNetReasonCode,")
        If sInsUpdAuthorizeNetReasonText.ToString = "'Select'" Then sb.Append("AuthorizeNetReasonText,")
        If sInsUpdAuthorizeNetApprovalCode.ToString = "'Select'" Then sb.Append("AuthorizeNetApprovalCode,")
        If sInsUpdAuthorizeNetAVSResultCode.ToString = "'Select'" Then sb.Append("AuthorizeNetAVSResultCode,")
        If sInsUpdAuthorizeNetAVSResultText.ToString = "'Select'" Then sb.Append("AuthorizeNetAVSResultText,")
        If sInsUpdAuthorizeNetTransactionID.ToString = "'Select'" Then sb.Append("AuthorizeNetTransactionID,")
        If sInsUpdAuthorizeNetCVCCResponseCode.ToString = "'Select'" Then sb.Append("AuthorizeNetCVCCResponseCode,")
        If sInsUpdAuthorizeNetCVCCResponseText.ToString = "'Select'" Then sb.Append("AuthorizeNetCVCCResponseText,")
        If sInsUpdLocation.ToString = "'Select'" Then sb.Append("Location,")
        If sInsUpdBankName.ToString = "'Select'" Then sb.Append("BankName,")
        If sInsUpdBankAccountType.ToString = "'Select'" Then sb.Append("BankAccountType,")
        If sInsUpdBankAccountName.ToString = "'Select'" Then sb.Append("BankAccountName,")
        If sInsUpdBankAccountNumber.ToString = "'Select'" Then sb.Append("BankAccountNumber,")
        If sInsUpdBankAccountNumberEncrypted.ToString = "'Select'" Then sb.Append("BankAccountNumberEncrypted,")
        If sInsUpdBankRoutingNumber.ToString = "'Select'" Then sb.Append("BankRoutingNumber,")
        If sInsUpdEmail.ToString = "'Select'" Then sb.Append("Email,")
        If sInsUpdQBActivityID.ToString = "'-12345'" Then sb.Append("QBActivityID,")
        If sInsUpdAuthorizeNetURL.ToString = "'Select'" Then sb.Append("AuthorizeNetURL,")
        If sInsUpdAuthNETPaymentProfileID.ToString = "'Select'" Then sb.Append("AuthNETPaymentProfileID,")
      Else
        sb.Append("PaymentID,")
        sb.Append("BookingID,")
        sb.Append("PropertyID,")
        sb.Append("HostID,")
        sb.Append("GuestID,")
        sb.Append("Category,")
        sb.Append("Type,")
        sb.Append("Amount,")
        sb.Append("[Date],")
        sb.Append("Computer,")
        sb.Append("CheckNumber,")
        sb.Append("CheckName,")
        sb.Append("CCType,")
        sb.Append("CCNumber,")
        sb.Append("CCNumberEncrypted,")
        sb.Append("CCName,")
        sb.Append("CCExpMonth,")
        sb.Append("CCExpYear,")
        sb.Append("CCVerification,")
        sb.Append("CCConfirm,")
        sb.Append("QBStatus,")
        sb.Append("Notes,")
        sb.Append("[User],")
        sb.Append("Status,")
        sb.Append("UseOnReceipt,")
        sb.Append("OrigCCNumber,")
        sb.Append("CCAddress,")
        sb.Append("CCCity,")
        sb.Append("CCState,")
        sb.Append("CCZip,")
        sb.Append("AuthorizeNetReturnCode,")
        sb.Append("AuthorizeNetReasonCode,")
        sb.Append("AuthorizeNetReasonText,")
        sb.Append("AuthorizeNetApprovalCode,")
        sb.Append("AuthorizeNetAVSResultCode,")
        sb.Append("AuthorizeNetAVSResultText,")
        sb.Append("AuthorizeNetTransactionID,")
        sb.Append("AuthorizeNetCVCCResponseCode,")
        sb.Append("AuthorizeNetCVCCResponseText,")
        sb.Append("Location,")
        sb.Append("BankName,")
        sb.Append("BankAccountType,")
        sb.Append("BankAccountName,")
        sb.Append("BankAccountNumber,")
        sb.Append("BankAccountNumberEncrypted,")
        sb.Append("BankRoutingNumber,")
        sb.Append("Email,")
        sb.Append("QBActivityID,")
        sb.Append("AuthorizeNetURL,")
        sb.Append("AuthNETPaymentProfileID,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [BalancePayments]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdPaymentID.ToString <> "") And (sInsUpdPaymentID <> "'-12345'") Then sbw.Append("PaymentID=" & sInsUpdPaymentID & " and ")
      If (sInsUpdBookingID.ToString <> "") And (sInsUpdBookingID <> "'-12345'") Then sbw.Append("BookingID=" & sInsUpdBookingID & " and ")
      If (sInsUpdPropertyID.ToString <> "") And (sInsUpdPropertyID <> "'-12345'") Then sbw.Append("PropertyID=" & sInsUpdPropertyID & " and ")
      If (sInsUpdHostID.ToString <> "") And (sInsUpdHostID <> "'-12345'") Then sbw.Append("HostID=" & sInsUpdHostID & " and ")
      If (sInsUpdGuestID.ToString <> "") And (sInsUpdGuestID <> "'-12345'") Then sbw.Append("GuestID=" & sInsUpdGuestID & " and ")
      If (sInsUpdCategory.ToString <> "") And (sInsUpdCategory <> "'Select'") Then sbw.Append("Category=" & sInsUpdCategory & " and ")
      If (sInsUpdType.ToString <> "") And (sInsUpdType <> "'Select'") Then sbw.Append("Type=" & sInsUpdType & " and ")
      If (sInsUpdAmount.ToString <> "") And (sInsUpdAmount <> "'-12345'") Then sbw.Append("Amount=" & sInsUpdAmount & " and ")
      If (sInsUpdDate.ToString <> "") And (sInsUpdDate <> "'Select'") Then sbw.Append("[Date]=" & sInsUpdDate & " and ")
      If (sInsUpdComputer.ToString <> "") And (sInsUpdComputer <> "'Select'") Then sbw.Append("Computer=" & sInsUpdComputer & " and ")
      If (sInsUpdCheckNumber.ToString <> "") And (sInsUpdCheckNumber <> "'Select'") Then sbw.Append("CheckNumber=" & sInsUpdCheckNumber & " and ")
      If (sInsUpdCheckName.ToString <> "") And (sInsUpdCheckName <> "'Select'") Then sbw.Append("CheckName=" & sInsUpdCheckName & " and ")
      If (sInsUpdCCType.ToString <> "") And (sInsUpdCCType <> "'Select'") Then sbw.Append("CCType=" & sInsUpdCCType & " and ")
      If (sInsUpdCCNumber.ToString <> "") And (sInsUpdCCNumber <> "'Select'") Then sbw.Append("CCNumber=" & sInsUpdCCNumber & " and ")
      If (sInsUpdCCNumberEncrypted.ToString <> "") And (sInsUpdCCNumberEncrypted <> "'Select'") Then sbw.Append("CCNumberEncrypted=" & sInsUpdCCNumberEncrypted & " and ")
      If (sInsUpdCCName.ToString <> "") And (sInsUpdCCName <> "'Select'") Then sbw.Append("CCName=" & sInsUpdCCName & " and ")
      If (sInsUpdCCExpMonth.ToString <> "") And (sInsUpdCCExpMonth <> "'Select'") Then sbw.Append("CCExpMonth=" & sInsUpdCCExpMonth & " and ")
      If (sInsUpdCCExpYear.ToString <> "") And (sInsUpdCCExpYear <> "'Select'") Then sbw.Append("CCExpYear=" & sInsUpdCCExpYear & " and ")
      If (sInsUpdCCVerification.ToString <> "") And (sInsUpdCCVerification <> "'Select'") Then sbw.Append("CCVerification=" & sInsUpdCCVerification & " and ")
      If (sInsUpdCCConfirm.ToString <> "") And (sInsUpdCCConfirm <> "'Select'") Then sbw.Append("CCConfirm=" & sInsUpdCCConfirm & " and ")
      If (sInsUpdQBStatus.ToString <> "") And (sInsUpdQBStatus <> "'Select'") Then sbw.Append("QBStatus=" & sInsUpdQBStatus & " and ")
      If (sInsUpdNotes.ToString <> "") And (sInsUpdNotes <> "'Select'") Then sbw.Append("Notes=" & sInsUpdNotes & " and ")
      If (sInsUpdUser.ToString <> "") And (sInsUpdUser <> "'Select'") Then sbw.Append("[User]=" & sInsUpdUser & " and ")
      If (sInsUpdStatus.ToString <> "") And (sInsUpdStatus <> "'Select'") Then sbw.Append("Status=" & sInsUpdStatus & " and ")
      If (sInsUpdUseOnReceipt.ToString <> "") And (sInsUpdUseOnReceipt <> "'-12345'") Then sbw.Append("UseOnReceipt=" & sInsUpdUseOnReceipt & " and ")
      If (sInsUpdOrigCCNumber.ToString <> "") And (sInsUpdOrigCCNumber <> "'Select'") Then sbw.Append("OrigCCNumber=" & sInsUpdOrigCCNumber & " and ")
      If (sInsUpdCCAddress.ToString <> "") And (sInsUpdCCAddress <> "'Select'") Then sbw.Append("CCAddress=" & sInsUpdCCAddress & " and ")
      If (sInsUpdCCCity.ToString <> "") And (sInsUpdCCCity <> "'Select'") Then sbw.Append("CCCity=" & sInsUpdCCCity & " and ")
      If (sInsUpdCCState.ToString <> "") And (sInsUpdCCState <> "'Select'") Then sbw.Append("CCState=" & sInsUpdCCState & " and ")
      If (sInsUpdCCZip.ToString <> "") And (sInsUpdCCZip <> "'Select'") Then sbw.Append("CCZip=" & sInsUpdCCZip & " and ")
      If (sInsUpdAuthorizeNetReturnCode.ToString <> "") And (sInsUpdAuthorizeNetReturnCode <> "'Select'") Then sbw.Append("AuthorizeNetReturnCode=" & sInsUpdAuthorizeNetReturnCode & " and ")
      If (sInsUpdAuthorizeNetReasonCode.ToString <> "") And (sInsUpdAuthorizeNetReasonCode <> "'Select'") Then sbw.Append("AuthorizeNetReasonCode=" & sInsUpdAuthorizeNetReasonCode & " and ")
      If (sInsUpdAuthorizeNetReasonText.ToString <> "") And (sInsUpdAuthorizeNetReasonText <> "'Select'") Then sbw.Append("AuthorizeNetReasonText=" & sInsUpdAuthorizeNetReasonText & " and ")
      If (sInsUpdAuthorizeNetApprovalCode.ToString <> "") And (sInsUpdAuthorizeNetApprovalCode <> "'Select'") Then sbw.Append("AuthorizeNetApprovalCode=" & sInsUpdAuthorizeNetApprovalCode & " and ")
      If (sInsUpdAuthorizeNetAVSResultCode.ToString <> "") And (sInsUpdAuthorizeNetAVSResultCode <> "'Select'") Then sbw.Append("AuthorizeNetAVSResultCode=" & sInsUpdAuthorizeNetAVSResultCode & " and ")
      If (sInsUpdAuthorizeNetAVSResultText.ToString <> "") And (sInsUpdAuthorizeNetAVSResultText <> "'Select'") Then sbw.Append("AuthorizeNetAVSResultText=" & sInsUpdAuthorizeNetAVSResultText & " and ")
      If (sInsUpdAuthorizeNetTransactionID.ToString <> "") And (sInsUpdAuthorizeNetTransactionID <> "'Select'") Then sbw.Append("AuthorizeNetTransactionID=" & sInsUpdAuthorizeNetTransactionID & " and ")
      If (sInsUpdAuthorizeNetCVCCResponseCode.ToString <> "") And (sInsUpdAuthorizeNetCVCCResponseCode <> "'Select'") Then sbw.Append("AuthorizeNetCVCCResponseCode=" & sInsUpdAuthorizeNetCVCCResponseCode & " and ")
      If (sInsUpdAuthorizeNetCVCCResponseText.ToString <> "") And (sInsUpdAuthorizeNetCVCCResponseText <> "'Select'") Then sbw.Append("AuthorizeNetCVCCResponseText=" & sInsUpdAuthorizeNetCVCCResponseText & " and ")
      If (sInsUpdLocation.ToString <> "") And (sInsUpdLocation <> "'Select'") Then sbw.Append("Location=" & sInsUpdLocation & " and ")
      If (sInsUpdBankName.ToString <> "") And (sInsUpdBankName <> "'Select'") Then sbw.Append("BankName=" & sInsUpdBankName & " and ")
      If (sInsUpdBankAccountType.ToString <> "") And (sInsUpdBankAccountType <> "'Select'") Then sbw.Append("BankAccountType=" & sInsUpdBankAccountType & " and ")
      If (sInsUpdBankAccountName.ToString <> "") And (sInsUpdBankAccountName <> "'Select'") Then sbw.Append("BankAccountName=" & sInsUpdBankAccountName & " and ")
      If (sInsUpdBankAccountNumber.ToString <> "") And (sInsUpdBankAccountNumber <> "'Select'") Then sbw.Append("BankAccountNumber=" & sInsUpdBankAccountNumber & " and ")
      If (sInsUpdBankAccountNumberEncrypted.ToString <> "") And (sInsUpdBankAccountNumberEncrypted <> "'Select'") Then sbw.Append("BankAccountNumberEncrypted=" & sInsUpdBankAccountNumberEncrypted & " and ")
      If (sInsUpdBankRoutingNumber.ToString <> "") And (sInsUpdBankRoutingNumber <> "'Select'") Then sbw.Append("BankRoutingNumber=" & sInsUpdBankRoutingNumber & " and ")
      If (sInsUpdEmail.ToString <> "") And (sInsUpdEmail <> "'Select'") Then sbw.Append("Email=" & sInsUpdEmail & " and ")
      If (sInsUpdQBActivityID.ToString <> "") And (sInsUpdQBActivityID <> "'-12345'") Then sbw.Append("QBActivityID=" & sInsUpdQBActivityID & " and ")
      If (sInsUpdAuthorizeNetURL.ToString <> "") And (sInsUpdAuthorizeNetURL <> "'Select'") Then sbw.Append("AuthorizeNetURL=" & sInsUpdAuthorizeNetURL & " and ")
      If (sInsUpdAuthNETPaymentProfileID.ToString <> "") And (sInsUpdAuthNETPaymentProfileID <> "'Select'") Then sbw.Append("AuthNETPaymentProfileID=" & sInsUpdAuthNETPaymentProfileID & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      PaymentID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("PaymentID")), 0, CurrentRow.Item("PaymentID").ToString)
      BookingID__Integer = IIf(IsDBNull(CurrentRow.Item("BookingID")), 0, CurrentRow.Item("BookingID"))
      PropertyID__Integer = IIf(IsDBNull(CurrentRow.Item("PropertyID")), 0, CurrentRow.Item("PropertyID"))
      HostID__Integer = IIf(IsDBNull(CurrentRow.Item("HostID")), 0, CurrentRow.Item("HostID"))
      GuestID__Integer = IIf(IsDBNull(CurrentRow.Item("GuestID")), 0, CurrentRow.Item("GuestID"))
      Category__String = IIf(IsDBNull(CurrentRow.Item("Category")), "", CurrentRow.Item("Category"))
      Type__String = IIf(IsDBNull(CurrentRow.Item("Type")), "", CurrentRow.Item("Type"))
      Amount__Numeric = IIf(IsDBNull(CurrentRow.Item("Amount")), 0.0, CurrentRow.Item("Amount"))
      Date__Date = IIf(IsDBNull(CurrentRow.Item("Date")), "", CurrentRow.Item("Date"))
      Computer__String = IIf(IsDBNull(CurrentRow.Item("Computer")), "", CurrentRow.Item("Computer"))
      CheckNumber__String = IIf(IsDBNull(CurrentRow.Item("CheckNumber")), "", CurrentRow.Item("CheckNumber"))
      CheckName__String = IIf(IsDBNull(CurrentRow.Item("CheckName")), "", CurrentRow.Item("CheckName"))
      CCType__String = IIf(IsDBNull(CurrentRow.Item("CCType")), "", CurrentRow.Item("CCType"))
      CCNumber__String = IIf(IsDBNull(CurrentRow.Item("CCNumber")), "", CurrentRow.Item("CCNumber"))
      CCNumberEncrypted__String = IIf(IsDBNull(CurrentRow.Item("CCNumberEncrypted")), "", CurrentRow.Item("CCNumberEncrypted"))
      CCName__String = IIf(IsDBNull(CurrentRow.Item("CCName")), "", CurrentRow.Item("CCName"))
      CCExpMonth__String = IIf(IsDBNull(CurrentRow.Item("CCExpMonth")), "", CurrentRow.Item("CCExpMonth"))
      CCExpYear__String = IIf(IsDBNull(CurrentRow.Item("CCExpYear")), "", CurrentRow.Item("CCExpYear"))
      CCVerification__String = IIf(IsDBNull(CurrentRow.Item("CCVerification")), "", CurrentRow.Item("CCVerification"))
      CCConfirm__String = IIf(IsDBNull(CurrentRow.Item("CCConfirm")), "", CurrentRow.Item("CCConfirm"))
      QBStatus__String = IIf(IsDBNull(CurrentRow.Item("QBStatus")), "", CurrentRow.Item("QBStatus"))
      Notes__String = IIf(IsDBNull(CurrentRow.Item("Notes")), "", CurrentRow.Item("Notes"))
      User__String = IIf(IsDBNull(CurrentRow.Item("User")), "", CurrentRow.Item("User"))
      Status__String = IIf(IsDBNull(CurrentRow.Item("Status")), "", CurrentRow.Item("Status"))
      UseOnReceipt__Integer = IIf(IsDBNull(CurrentRow.Item("UseOnReceipt")), 0, CurrentRow.Item("UseOnReceipt"))
      OrigCCNumber__String = IIf(IsDBNull(CurrentRow.Item("OrigCCNumber")), "", CurrentRow.Item("OrigCCNumber"))
      CCAddress__String = IIf(IsDBNull(CurrentRow.Item("CCAddress")), "", CurrentRow.Item("CCAddress"))
      CCCity__String = IIf(IsDBNull(CurrentRow.Item("CCCity")), "", CurrentRow.Item("CCCity"))
      CCState__String = IIf(IsDBNull(CurrentRow.Item("CCState")), "", CurrentRow.Item("CCState"))
      CCZip__String = IIf(IsDBNull(CurrentRow.Item("CCZip")), "", CurrentRow.Item("CCZip"))
      AuthorizeNetReturnCode__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetReturnCode")), "", CurrentRow.Item("AuthorizeNetReturnCode"))
      AuthorizeNetReasonCode__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetReasonCode")), "", CurrentRow.Item("AuthorizeNetReasonCode"))
      AuthorizeNetReasonText__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetReasonText")), "", CurrentRow.Item("AuthorizeNetReasonText"))
      AuthorizeNetApprovalCode__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetApprovalCode")), "", CurrentRow.Item("AuthorizeNetApprovalCode"))
      AuthorizeNetAVSResultCode__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetAVSResultCode")), "", CurrentRow.Item("AuthorizeNetAVSResultCode"))
      AuthorizeNetAVSResultText__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetAVSResultText")), "", CurrentRow.Item("AuthorizeNetAVSResultText"))
      AuthorizeNetTransactionID__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetTransactionID")), "", CurrentRow.Item("AuthorizeNetTransactionID"))
      AuthorizeNetCVCCResponseCode__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetCVCCResponseCode")), "", CurrentRow.Item("AuthorizeNetCVCCResponseCode"))
      AuthorizeNetCVCCResponseText__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetCVCCResponseText")), "", CurrentRow.Item("AuthorizeNetCVCCResponseText"))
      Location__String = IIf(IsDBNull(CurrentRow.Item("Location")), "", CurrentRow.Item("Location"))
      BankName__String = IIf(IsDBNull(CurrentRow.Item("BankName")), "", CurrentRow.Item("BankName"))
      BankAccountType__String = IIf(IsDBNull(CurrentRow.Item("BankAccountType")), "", CurrentRow.Item("BankAccountType"))
      BankAccountName__String = IIf(IsDBNull(CurrentRow.Item("BankAccountName")), "", CurrentRow.Item("BankAccountName"))
      BankAccountNumber__String = IIf(IsDBNull(CurrentRow.Item("BankAccountNumber")), "", CurrentRow.Item("BankAccountNumber"))
      BankAccountNumberEncrypted__String = IIf(IsDBNull(CurrentRow.Item("BankAccountNumberEncrypted")), "", CurrentRow.Item("BankAccountNumberEncrypted"))
      BankRoutingNumber__String = IIf(IsDBNull(CurrentRow.Item("BankRoutingNumber")), "", CurrentRow.Item("BankRoutingNumber"))
      Email__String = IIf(IsDBNull(CurrentRow.Item("Email")), "", CurrentRow.Item("Email"))
      QBActivityID__Integer = IIf(IsDBNull(CurrentRow.Item("QBActivityID")), 0, CurrentRow.Item("QBActivityID"))
      AuthorizeNetURL__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetURL")), "", CurrentRow.Item("AuthorizeNetURL"))
      AuthNETPaymentProfileID__String = IIf(IsDBNull(CurrentRow.Item("AuthNETPaymentProfileID")), "", CurrentRow.Item("AuthNETPaymentProfileID"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [BalancePayments](")
    If sInsUpdBookingID.ToString <> "" Then
      sb.Append("BookingID,")
      sbv.Append(sInsUpdBookingID & ",")
    End If
    If sInsUpdPropertyID.ToString <> "" Then
      sb.Append("PropertyID,")
      sbv.Append(sInsUpdPropertyID & ",")
    End If
    If sInsUpdHostID.ToString <> "" Then
      sb.Append("HostID,")
      sbv.Append(sInsUpdHostID & ",")
    End If
    If sInsUpdGuestID.ToString <> "" Then
      sb.Append("GuestID,")
      sbv.Append(sInsUpdGuestID & ",")
    End If
    If sInsUpdCategory.ToString <> "" Then
      sb.Append("Category,")
      sbv.Append(sInsUpdCategory & ",")
    End If
    If sInsUpdType.ToString <> "" Then
      sb.Append("Type,")
      sbv.Append(sInsUpdType & ",")
    End If
    If sInsUpdAmount.ToString <> "" Then
      sb.Append("Amount,")
      sbv.Append(sInsUpdAmount & ",")
    End If
    If sInsUpdDate.ToString <> "" Then
      sb.Append("[Date],")
      sbv.Append(sInsUpdDate & ",")
    End If
    If sInsUpdComputer.ToString <> "" Then
      sb.Append("Computer,")
      sbv.Append(sInsUpdComputer & ",")
    End If
    If sInsUpdCheckNumber.ToString <> "" Then
      sb.Append("CheckNumber,")
      sbv.Append(sInsUpdCheckNumber & ",")
    End If
    If sInsUpdCheckName.ToString <> "" Then
      sb.Append("CheckName,")
      sbv.Append(sInsUpdCheckName & ",")
    End If
    If sInsUpdCCType.ToString <> "" Then
      sb.Append("CCType,")
      sbv.Append(sInsUpdCCType & ",")
    End If
    If sInsUpdCCNumber.ToString <> "" Then
      sb.Append("CCNumber,")
      sbv.Append(sInsUpdCCNumber & ",")
    End If
    If sInsUpdCCNumberEncrypted.ToString <> "" Then
      sb.Append("CCNumberEncrypted,")
      sbv.Append(sInsUpdCCNumberEncrypted & ",")
    End If
    If sInsUpdCCName.ToString <> "" Then
      sb.Append("CCName,")
      sbv.Append(sInsUpdCCName & ",")
    End If
    If sInsUpdCCExpMonth.ToString <> "" Then
      sb.Append("CCExpMonth,")
      sbv.Append(sInsUpdCCExpMonth & ",")
    End If
    If sInsUpdCCExpYear.ToString <> "" Then
      sb.Append("CCExpYear,")
      sbv.Append(sInsUpdCCExpYear & ",")
    End If
    If sInsUpdCCVerification.ToString <> "" Then
      sb.Append("CCVerification,")
      sbv.Append(sInsUpdCCVerification & ",")
    End If
    If sInsUpdCCConfirm.ToString <> "" Then
      sb.Append("CCConfirm,")
      sbv.Append(sInsUpdCCConfirm & ",")
    End If
    If sInsUpdQBStatus.ToString <> "" Then
      sb.Append("QBStatus,")
      sbv.Append(sInsUpdQBStatus & ",")
    End If
    If sInsUpdNotes.ToString <> "" Then
      sb.Append("Notes,")
      sbv.Append(sInsUpdNotes & ",")
    End If
    If sInsUpdUser.ToString <> "" Then
      sb.Append("[User],")
      sbv.Append(sInsUpdUser & ",")
    End If
    If sInsUpdStatus.ToString <> "" Then
      sb.Append("Status,")
      sbv.Append(sInsUpdStatus & ",")
    End If
    If sInsUpdUseOnReceipt.ToString <> "" Then
      sb.Append("UseOnReceipt,")
      sbv.Append(sInsUpdUseOnReceipt & ",")
    End If
    If sInsUpdOrigCCNumber.ToString <> "" Then
      sb.Append("OrigCCNumber,")
      sbv.Append(sInsUpdOrigCCNumber & ",")
    End If
    If sInsUpdCCAddress.ToString <> "" Then
      sb.Append("CCAddress,")
      sbv.Append(sInsUpdCCAddress & ",")
    End If
    If sInsUpdCCCity.ToString <> "" Then
      sb.Append("CCCity,")
      sbv.Append(sInsUpdCCCity & ",")
    End If
    If sInsUpdCCState.ToString <> "" Then
      sb.Append("CCState,")
      sbv.Append(sInsUpdCCState & ",")
    End If
    If sInsUpdCCZip.ToString <> "" Then
      sb.Append("CCZip,")
      sbv.Append(sInsUpdCCZip & ",")
    End If
    If sInsUpdAuthorizeNetReturnCode.ToString <> "" Then
      sb.Append("AuthorizeNetReturnCode,")
      sbv.Append(sInsUpdAuthorizeNetReturnCode & ",")
    End If
    If sInsUpdAuthorizeNetReasonCode.ToString <> "" Then
      sb.Append("AuthorizeNetReasonCode,")
      sbv.Append(sInsUpdAuthorizeNetReasonCode & ",")
    End If
    If sInsUpdAuthorizeNetReasonText.ToString <> "" Then
      sb.Append("AuthorizeNetReasonText,")
      sbv.Append(sInsUpdAuthorizeNetReasonText & ",")
    End If
    If sInsUpdAuthorizeNetApprovalCode.ToString <> "" Then
      sb.Append("AuthorizeNetApprovalCode,")
      sbv.Append(sInsUpdAuthorizeNetApprovalCode & ",")
    End If
    If sInsUpdAuthorizeNetAVSResultCode.ToString <> "" Then
      sb.Append("AuthorizeNetAVSResultCode,")
      sbv.Append(sInsUpdAuthorizeNetAVSResultCode & ",")
    End If
    If sInsUpdAuthorizeNetAVSResultText.ToString <> "" Then
      sb.Append("AuthorizeNetAVSResultText,")
      sbv.Append(sInsUpdAuthorizeNetAVSResultText & ",")
    End If
    If sInsUpdAuthorizeNetTransactionID.ToString <> "" Then
      sb.Append("AuthorizeNetTransactionID,")
      sbv.Append(sInsUpdAuthorizeNetTransactionID & ",")
    End If
    If sInsUpdAuthorizeNetCVCCResponseCode.ToString <> "" Then
      sb.Append("AuthorizeNetCVCCResponseCode,")
      sbv.Append(sInsUpdAuthorizeNetCVCCResponseCode & ",")
    End If
    If sInsUpdAuthorizeNetCVCCResponseText.ToString <> "" Then
      sb.Append("AuthorizeNetCVCCResponseText,")
      sbv.Append(sInsUpdAuthorizeNetCVCCResponseText & ",")
    End If
    If sInsUpdLocation.ToString <> "" Then
      sb.Append("Location,")
      sbv.Append(sInsUpdLocation & ",")
    End If
    If sInsUpdBankName.ToString <> "" Then
      sb.Append("BankName,")
      sbv.Append(sInsUpdBankName & ",")
    End If
    If sInsUpdBankAccountType.ToString <> "" Then
      sb.Append("BankAccountType,")
      sbv.Append(sInsUpdBankAccountType & ",")
    End If
    If sInsUpdBankAccountName.ToString <> "" Then
      sb.Append("BankAccountName,")
      sbv.Append(sInsUpdBankAccountName & ",")
    End If
    If sInsUpdBankAccountNumber.ToString <> "" Then
      sb.Append("BankAccountNumber,")
      sbv.Append(sInsUpdBankAccountNumber & ",")
    End If
    If sInsUpdBankAccountNumberEncrypted.ToString <> "" Then
      sb.Append("BankAccountNumberEncrypted,")
      sbv.Append(sInsUpdBankAccountNumberEncrypted & ",")
    End If
    If sInsUpdBankRoutingNumber.ToString <> "" Then
      sb.Append("BankRoutingNumber,")
      sbv.Append(sInsUpdBankRoutingNumber & ",")
    End If
    If sInsUpdEmail.ToString <> "" Then
      sb.Append("Email,")
      sbv.Append(sInsUpdEmail & ",")
    End If
    If sInsUpdQBActivityID.ToString <> "" Then
      sb.Append("QBActivityID,")
      sbv.Append(sInsUpdQBActivityID & ",")
    End If
    If sInsUpdAuthorizeNetURL.ToString <> "" Then
      sb.Append("AuthorizeNetURL,")
      sbv.Append(sInsUpdAuthorizeNetURL & ",")
    End If
    If sInsUpdAuthNETPaymentProfileID.ToString <> "" Then
      sb.Append("AuthNETPaymentProfileID,")
      sbv.Append(sInsUpdAuthNETPaymentProfileID & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(PaymentID) from [BalancePayments]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    PaymentID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [BalancePayments] Set ")
    If sInsUpdBookingID.ToString <> "" Then sb.Append("BookingID=" & sInsUpdBookingID & ",")
    If sInsUpdPropertyID.ToString <> "" Then sb.Append("PropertyID=" & sInsUpdPropertyID & ",")
    If sInsUpdHostID.ToString <> "" Then sb.Append("HostID=" & sInsUpdHostID & ",")
    If sInsUpdGuestID.ToString <> "" Then sb.Append("GuestID=" & sInsUpdGuestID & ",")
    If sInsUpdCategory.ToString <> "" Then sb.Append("Category=" & sInsUpdCategory & ",")
    If sInsUpdType.ToString <> "" Then sb.Append("Type=" & sInsUpdType & ",")
    If sInsUpdAmount.ToString <> "" Then sb.Append("Amount=" & sInsUpdAmount & ",")
    If sInsUpdDate.ToString <> "" Then sb.Append("[Date]=" & sInsUpdDate & ",")
    If sInsUpdComputer.ToString <> "" Then sb.Append("Computer=" & sInsUpdComputer & ",")
    If sInsUpdCheckNumber.ToString <> "" Then sb.Append("CheckNumber=" & sInsUpdCheckNumber & ",")
    If sInsUpdCheckName.ToString <> "" Then sb.Append("CheckName=" & sInsUpdCheckName & ",")
    If sInsUpdCCType.ToString <> "" Then sb.Append("CCType=" & sInsUpdCCType & ",")
    If sInsUpdCCNumber.ToString <> "" Then sb.Append("CCNumber=" & sInsUpdCCNumber & ",")
    If sInsUpdCCNumberEncrypted.ToString <> "" Then sb.Append("CCNumberEncrypted=" & sInsUpdCCNumberEncrypted & ",")
    If sInsUpdCCName.ToString <> "" Then sb.Append("CCName=" & sInsUpdCCName & ",")
    If sInsUpdCCExpMonth.ToString <> "" Then sb.Append("CCExpMonth=" & sInsUpdCCExpMonth & ",")
    If sInsUpdCCExpYear.ToString <> "" Then sb.Append("CCExpYear=" & sInsUpdCCExpYear & ",")
    If sInsUpdCCVerification.ToString <> "" Then sb.Append("CCVerification=" & sInsUpdCCVerification & ",")
    If sInsUpdCCConfirm.ToString <> "" Then sb.Append("CCConfirm=" & sInsUpdCCConfirm & ",")
    If sInsUpdQBStatus.ToString <> "" Then sb.Append("QBStatus=" & sInsUpdQBStatus & ",")
    If sInsUpdNotes.ToString <> "" Then sb.Append("Notes=" & sInsUpdNotes & ",")
    If sInsUpdUser.ToString <> "" Then sb.Append("[User]=" & sInsUpdUser & ",")
    If sInsUpdStatus.ToString <> "" Then sb.Append("Status=" & sInsUpdStatus & ",")
    If sInsUpdUseOnReceipt.ToString <> "" Then sb.Append("UseOnReceipt=" & sInsUpdUseOnReceipt & ",")
    If sInsUpdOrigCCNumber.ToString <> "" Then sb.Append("OrigCCNumber=" & sInsUpdOrigCCNumber & ",")
    If sInsUpdCCAddress.ToString <> "" Then sb.Append("CCAddress=" & sInsUpdCCAddress & ",")
    If sInsUpdCCCity.ToString <> "" Then sb.Append("CCCity=" & sInsUpdCCCity & ",")
    If sInsUpdCCState.ToString <> "" Then sb.Append("CCState=" & sInsUpdCCState & ",")
    If sInsUpdCCZip.ToString <> "" Then sb.Append("CCZip=" & sInsUpdCCZip & ",")
    If sInsUpdAuthorizeNetReturnCode.ToString <> "" Then sb.Append("AuthorizeNetReturnCode=" & sInsUpdAuthorizeNetReturnCode & ",")
    If sInsUpdAuthorizeNetReasonCode.ToString <> "" Then sb.Append("AuthorizeNetReasonCode=" & sInsUpdAuthorizeNetReasonCode & ",")
    If sInsUpdAuthorizeNetReasonText.ToString <> "" Then sb.Append("AuthorizeNetReasonText=" & sInsUpdAuthorizeNetReasonText & ",")
    If sInsUpdAuthorizeNetApprovalCode.ToString <> "" Then sb.Append("AuthorizeNetApprovalCode=" & sInsUpdAuthorizeNetApprovalCode & ",")
    If sInsUpdAuthorizeNetAVSResultCode.ToString <> "" Then sb.Append("AuthorizeNetAVSResultCode=" & sInsUpdAuthorizeNetAVSResultCode & ",")
    If sInsUpdAuthorizeNetAVSResultText.ToString <> "" Then sb.Append("AuthorizeNetAVSResultText=" & sInsUpdAuthorizeNetAVSResultText & ",")
    If sInsUpdAuthorizeNetTransactionID.ToString <> "" Then sb.Append("AuthorizeNetTransactionID=" & sInsUpdAuthorizeNetTransactionID & ",")
    If sInsUpdAuthorizeNetCVCCResponseCode.ToString <> "" Then sb.Append("AuthorizeNetCVCCResponseCode=" & sInsUpdAuthorizeNetCVCCResponseCode & ",")
    If sInsUpdAuthorizeNetCVCCResponseText.ToString <> "" Then sb.Append("AuthorizeNetCVCCResponseText=" & sInsUpdAuthorizeNetCVCCResponseText & ",")
    If sInsUpdLocation.ToString <> "" Then sb.Append("Location=" & sInsUpdLocation & ",")
    If sInsUpdBankName.ToString <> "" Then sb.Append("BankName=" & sInsUpdBankName & ",")
    If sInsUpdBankAccountType.ToString <> "" Then sb.Append("BankAccountType=" & sInsUpdBankAccountType & ",")
    If sInsUpdBankAccountName.ToString <> "" Then sb.Append("BankAccountName=" & sInsUpdBankAccountName & ",")
    If sInsUpdBankAccountNumber.ToString <> "" Then sb.Append("BankAccountNumber=" & sInsUpdBankAccountNumber & ",")
    If sInsUpdBankAccountNumberEncrypted.ToString <> "" Then sb.Append("BankAccountNumberEncrypted=" & sInsUpdBankAccountNumberEncrypted & ",")
    If sInsUpdBankRoutingNumber.ToString <> "" Then sb.Append("BankRoutingNumber=" & sInsUpdBankRoutingNumber & ",")
    If sInsUpdEmail.ToString <> "" Then sb.Append("Email=" & sInsUpdEmail & ",")
    If sInsUpdQBActivityID.ToString <> "" Then sb.Append("QBActivityID=" & sInsUpdQBActivityID & ",")
    If sInsUpdAuthorizeNetURL.ToString <> "" Then sb.Append("AuthorizeNetURL=" & sInsUpdAuthorizeNetURL & ",")
    If sInsUpdAuthNETPaymentProfileID.ToString <> "" Then sb.Append("AuthNETPaymentProfileID=" & sInsUpdAuthNETPaymentProfileID & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where PaymentID=" & sInsUpdPaymentID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [BalancePayments] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("PaymentID=" & sInsUpdPaymentID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TableDepositPayments

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iPaymentID As Int32
  Private sInsUpdPaymentID As String
  Property PaymentID_PK__Integer() As Int32
    Get
      Return iPaymentID
    End Get
    Set(ByVal Value As Int32)
      iPaymentID = Value
      sInsUpdPaymentID = oUtil.FixParam(iPaymentID, True)
    End Set
  End Property

  Private iBookingID As Int32
  Private sInsUpdBookingID As String
  Property BookingID__Integer() As Int32
    Get
      Return iBookingID
    End Get
    Set(ByVal Value As Int32)
      iBookingID = Value
      sInsUpdBookingID = oUtil.FixParam(iBookingID, True)
    End Set
  End Property

  Private iPropertyID As Int32
  Private sInsUpdPropertyID As String
  Property PropertyID__Integer() As Int32
    Get
      Return iPropertyID
    End Get
    Set(ByVal Value As Int32)
      iPropertyID = Value
      sInsUpdPropertyID = oUtil.FixParam(iPropertyID, True)
    End Set
  End Property

  Private iHostID As Int32
  Private sInsUpdHostID As String
  Property HostID__Integer() As Int32
    Get
      Return iHostID
    End Get
    Set(ByVal Value As Int32)
      iHostID = Value
      sInsUpdHostID = oUtil.FixParam(iHostID, True)
    End Set
  End Property

  Private iGuestID As Int32
  Private sInsUpdGuestID As String
  Property GuestID__Integer() As Int32
    Get
      Return iGuestID
    End Get
    Set(ByVal Value As Int32)
      iGuestID = Value
      sInsUpdGuestID = oUtil.FixParam(iGuestID, True)
    End Set
  End Property

  Private sCategory As String
  Private sInsUpdCategory As String
  Property Category__String() As String
    Get
      Return sCategory
    End Get
    Set(ByVal Value As String)
      sCategory = Value
      sInsUpdCategory = oUtil.FixParam(sCategory, True)
    End Set
  End Property

  Private sType As String
  Private sInsUpdType As String
  Property Type__String() As String
    Get
      Return sType
    End Get
    Set(ByVal Value As String)
      sType = Value
      sInsUpdType = oUtil.FixParam(sType, True)
    End Set
  End Property

  Private dAmount As Double
  Private sInsUpdAmount As String
  Property Amount__Numeric() As Double
    Get
      Return dAmount
    End Get
    Set(ByVal Value As Double)
      dAmount = Value
      sInsUpdAmount = oUtil.FixParam(dAmount, True)
    End Set
  End Property

  Private sDate As String
  Private sInsUpdDate As String
  Property Date__Date() As String
    Get
      Return sDate
    End Get
    Set(ByVal Value As String)
      sDate = Value
      sInsUpdDate = oUtil.FixParam(sDate, True)
    End Set
  End Property

  Private sComputer As String
  Private sInsUpdComputer As String
  Property Computer__String() As String
    Get
      Return sComputer
    End Get
    Set(ByVal Value As String)
      sComputer = Value
      sInsUpdComputer = oUtil.FixParam(sComputer, True)
    End Set
  End Property

  Private sCheckNumber As String
  Private sInsUpdCheckNumber As String
  Property CheckNumber__String() As String
    Get
      Return sCheckNumber
    End Get
    Set(ByVal Value As String)
      sCheckNumber = Value
      sInsUpdCheckNumber = oUtil.FixParam(sCheckNumber, True)
    End Set
  End Property

  Private sCheckName As String
  Private sInsUpdCheckName As String
  Property CheckName__String() As String
    Get
      Return sCheckName
    End Get
    Set(ByVal Value As String)
      sCheckName = Value
      sInsUpdCheckName = oUtil.FixParam(sCheckName, True)
    End Set
  End Property

  Private sCCType As String
  Private sInsUpdCCType As String
  Property CCType__String() As String
    Get
      Return sCCType
    End Get
    Set(ByVal Value As String)
      sCCType = Value
      sInsUpdCCType = oUtil.FixParam(sCCType, True)
    End Set
  End Property

  Private sCCNumber As String
  Private sInsUpdCCNumber As String
  Property CCNumber__String() As String
    Get
      Return sCCNumber
    End Get
    Set(ByVal Value As String)
      sCCNumber = Value
      sInsUpdCCNumber = oUtil.FixParam(sCCNumber, True)
    End Set
  End Property

  Private sCCNumberEncrypted As String
  Private sInsUpdCCNumberEncrypted As String
  Property CCNumberEncrypted__String() As String
    Get
      Return sCCNumberEncrypted
    End Get
    Set(ByVal Value As String)
      sCCNumberEncrypted = Value
      sInsUpdCCNumberEncrypted = oUtil.FixParam(sCCNumberEncrypted, True)
    End Set
  End Property

  Private sCCName As String
  Private sInsUpdCCName As String
  Property CCName__String() As String
    Get
      Return sCCName
    End Get
    Set(ByVal Value As String)
      sCCName = Value
      sInsUpdCCName = oUtil.FixParam(sCCName, True)
    End Set
  End Property

  Private sCCExpMonth As String
  Private sInsUpdCCExpMonth As String
  Property CCExpMonth__String() As String
    Get
      Return sCCExpMonth
    End Get
    Set(ByVal Value As String)
      sCCExpMonth = Value
      sInsUpdCCExpMonth = oUtil.FixParam(sCCExpMonth, True)
    End Set
  End Property

  Private sCCExpYear As String
  Private sInsUpdCCExpYear As String
  Property CCExpYear__String() As String
    Get
      Return sCCExpYear
    End Get
    Set(ByVal Value As String)
      sCCExpYear = Value
      sInsUpdCCExpYear = oUtil.FixParam(sCCExpYear, True)
    End Set
  End Property

  Private sCCVerification As String
  Private sInsUpdCCVerification As String
  Property CCVerification__String() As String
    Get
      Return sCCVerification
    End Get
    Set(ByVal Value As String)
      sCCVerification = Value
      sInsUpdCCVerification = oUtil.FixParam(sCCVerification, True)
    End Set
  End Property

  Private sCCConfirm As String
  Private sInsUpdCCConfirm As String
  Property CCConfirm__String() As String
    Get
      Return sCCConfirm
    End Get
    Set(ByVal Value As String)
      sCCConfirm = Value
      sInsUpdCCConfirm = oUtil.FixParam(sCCConfirm, True)
    End Set
  End Property

  Private sQBStatus As String
  Private sInsUpdQBStatus As String
  Property QBStatus__String() As String
    Get
      Return sQBStatus
    End Get
    Set(ByVal Value As String)
      sQBStatus = Value
      sInsUpdQBStatus = oUtil.FixParam(sQBStatus, True)
    End Set
  End Property

  Private sNotes As String
  Private sInsUpdNotes As String
  Property Notes__String() As String
    Get
      Return sNotes
    End Get
    Set(ByVal Value As String)
      sNotes = Value
      sInsUpdNotes = oUtil.FixParam(sNotes, True)
    End Set
  End Property

  Private sUser As String
  Private sInsUpdUser As String
  Property User__String() As String
    Get
      Return sUser
    End Get
    Set(ByVal Value As String)
      sUser = Value
      sInsUpdUser = oUtil.FixParam(sUser, True)
    End Set
  End Property

  Private sStatus As String
  Private sInsUpdStatus As String
  Property Status__String() As String
    Get
      Return sStatus
    End Get
    Set(ByVal Value As String)
      sStatus = Value
      sInsUpdStatus = oUtil.FixParam(sStatus, True)
    End Set
  End Property

  Private iUseOnReceipt As Int32
  Private sInsUpdUseOnReceipt As String
  Property UseOnReceipt__Integer() As Int32
    Get
      Return iUseOnReceipt
    End Get
    Set(ByVal Value As Int32)
      iUseOnReceipt = Value
      sInsUpdUseOnReceipt = oUtil.FixParam(iUseOnReceipt, True)
    End Set
  End Property

  Private sOrigCCNumber As String
  Private sInsUpdOrigCCNumber As String
  Property OrigCCNumber__String() As String
    Get
      Return sOrigCCNumber
    End Get
    Set(ByVal Value As String)
      sOrigCCNumber = Value
      sInsUpdOrigCCNumber = oUtil.FixParam(sOrigCCNumber, True)
    End Set
  End Property

  Private sCCAddress As String
  Private sInsUpdCCAddress As String
  Property CCAddress__String() As String
    Get
      Return sCCAddress
    End Get
    Set(ByVal Value As String)
      sCCAddress = Value
      sInsUpdCCAddress = oUtil.FixParam(sCCAddress, True)
    End Set
  End Property

  Private sCCCity As String
  Private sInsUpdCCCity As String
  Property CCCity__String() As String
    Get
      Return sCCCity
    End Get
    Set(ByVal Value As String)
      sCCCity = Value
      sInsUpdCCCity = oUtil.FixParam(sCCCity, True)
    End Set
  End Property

  Private sCCState As String
  Private sInsUpdCCState As String
  Property CCState__String() As String
    Get
      Return sCCState
    End Get
    Set(ByVal Value As String)
      sCCState = Value
      sInsUpdCCState = oUtil.FixParam(sCCState, True)
    End Set
  End Property

  Private sCCZip As String
  Private sInsUpdCCZip As String
  Property CCZip__String() As String
    Get
      Return sCCZip
    End Get
    Set(ByVal Value As String)
      sCCZip = Value
      sInsUpdCCZip = oUtil.FixParam(sCCZip, True)
    End Set
  End Property

  Private sCCCountry As String
  Private sInsUpdCCCountry As String
  Property CCCountry__String() As String
    Get
      Return sCCCountry
    End Get
    Set(ByVal Value As String)
      sCCCountry = Value
      sInsUpdCCCountry = oUtil.FixParam(sCCCountry, True)
    End Set
  End Property

  Private sAuthorizeNetReturnCode As String
  Private sInsUpdAuthorizeNetReturnCode As String
  Property AuthorizeNetReturnCode__String() As String
    Get
      Return sAuthorizeNetReturnCode
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetReturnCode = Value
      sInsUpdAuthorizeNetReturnCode = oUtil.FixParam(sAuthorizeNetReturnCode, True)
    End Set
  End Property

  Private sAuthorizeNetReasonCode As String
  Private sInsUpdAuthorizeNetReasonCode As String
  Property AuthorizeNetReasonCode__String() As String
    Get
      Return sAuthorizeNetReasonCode
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetReasonCode = Value
      sInsUpdAuthorizeNetReasonCode = oUtil.FixParam(sAuthorizeNetReasonCode, True)
    End Set
  End Property

  Private sAuthorizeNetReasonText As String
  Private sInsUpdAuthorizeNetReasonText As String
  Property AuthorizeNetReasonText__String() As String
    Get
      Return sAuthorizeNetReasonText
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetReasonText = Value
      sInsUpdAuthorizeNetReasonText = oUtil.FixParam(sAuthorizeNetReasonText, True)
    End Set
  End Property

  Private sAuthorizeNetApprovalCode As String
  Private sInsUpdAuthorizeNetApprovalCode As String
  Property AuthorizeNetApprovalCode__String() As String
    Get
      Return sAuthorizeNetApprovalCode
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetApprovalCode = Value
      sInsUpdAuthorizeNetApprovalCode = oUtil.FixParam(sAuthorizeNetApprovalCode, True)
    End Set
  End Property

  Private sAuthorizeNetAVSResultCode As String
  Private sInsUpdAuthorizeNetAVSResultCode As String
  Property AuthorizeNetAVSResultCode__String() As String
    Get
      Return sAuthorizeNetAVSResultCode
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetAVSResultCode = Value
      sInsUpdAuthorizeNetAVSResultCode = oUtil.FixParam(sAuthorizeNetAVSResultCode, True)
    End Set
  End Property

  Private sAuthorizeNetAVSResultText As String
  Private sInsUpdAuthorizeNetAVSResultText As String
  Property AuthorizeNetAVSResultText__String() As String
    Get
      Return sAuthorizeNetAVSResultText
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetAVSResultText = Value
      sInsUpdAuthorizeNetAVSResultText = oUtil.FixParam(sAuthorizeNetAVSResultText, True)
    End Set
  End Property

  Private sAuthorizeNetTransactionID As String
  Private sInsUpdAuthorizeNetTransactionID As String
  Property AuthorizeNetTransactionID__String() As String
    Get
      Return sAuthorizeNetTransactionID
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetTransactionID = Value
      sInsUpdAuthorizeNetTransactionID = oUtil.FixParam(sAuthorizeNetTransactionID, True)
    End Set
  End Property

  Private sAuthorizeNetCVCCResponseCode As String
  Private sInsUpdAuthorizeNetCVCCResponseCode As String
  Property AuthorizeNetCVCCResponseCode__String() As String
    Get
      Return sAuthorizeNetCVCCResponseCode
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetCVCCResponseCode = Value
      sInsUpdAuthorizeNetCVCCResponseCode = oUtil.FixParam(sAuthorizeNetCVCCResponseCode, True)
    End Set
  End Property

  Private sAuthorizeNetCVCCResponseText As String
  Private sInsUpdAuthorizeNetCVCCResponseText As String
  Property AuthorizeNetCVCCResponseText__String() As String
    Get
      Return sAuthorizeNetCVCCResponseText
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetCVCCResponseText = Value
      sInsUpdAuthorizeNetCVCCResponseText = oUtil.FixParam(sAuthorizeNetCVCCResponseText, True)
    End Set
  End Property

  Private sLocation As String
  Private sInsUpdLocation As String
  Property Location__String() As String
    Get
      Return sLocation
    End Get
    Set(ByVal Value As String)
      sLocation = Value
      sInsUpdLocation = oUtil.FixParam(sLocation, True)
    End Set
  End Property

  Private sBankName As String
  Private sInsUpdBankName As String
  Property BankName__String() As String
    Get
      Return sBankName
    End Get
    Set(ByVal Value As String)
      sBankName = Value
      sInsUpdBankName = oUtil.FixParam(sBankName, True)
    End Set
  End Property

  Private sBankAccountType As String
  Private sInsUpdBankAccountType As String
  Property BankAccountType__String() As String
    Get
      Return sBankAccountType
    End Get
    Set(ByVal Value As String)
      sBankAccountType = Value
      sInsUpdBankAccountType = oUtil.FixParam(sBankAccountType, True)
    End Set
  End Property

  Private sBankAccountName As String
  Private sInsUpdBankAccountName As String
  Property BankAccountName__String() As String
    Get
      Return sBankAccountName
    End Get
    Set(ByVal Value As String)
      sBankAccountName = Value
      sInsUpdBankAccountName = oUtil.FixParam(sBankAccountName, True)
    End Set
  End Property

  Private sBankAccountNumber As String
  Private sInsUpdBankAccountNumber As String
  Property BankAccountNumber__String() As String
    Get
      Return sBankAccountNumber
    End Get
    Set(ByVal Value As String)
      sBankAccountNumber = Value
      sInsUpdBankAccountNumber = oUtil.FixParam(sBankAccountNumber, True)
    End Set
  End Property

  Private sBankAccountNumberEncrypted As String
  Private sInsUpdBankAccountNumberEncrypted As String
  Property BankAccountNumberEncrypted__String() As String
    Get
      Return sBankAccountNumberEncrypted
    End Get
    Set(ByVal Value As String)
      sBankAccountNumberEncrypted = Value
      sInsUpdBankAccountNumberEncrypted = oUtil.FixParam(sBankAccountNumberEncrypted, True)
    End Set
  End Property

  Private sBankRoutingNumber As String
  Private sInsUpdBankRoutingNumber As String
  Property BankRoutingNumber__String() As String
    Get
      Return sBankRoutingNumber
    End Get
    Set(ByVal Value As String)
      sBankRoutingNumber = Value
      sInsUpdBankRoutingNumber = oUtil.FixParam(sBankRoutingNumber, True)
    End Set
  End Property

  Private sEmail As String
  Private sInsUpdEmail As String
  Property Email__String() As String
    Get
      Return sEmail
    End Get
    Set(ByVal Value As String)
      sEmail = Value
      sInsUpdEmail = oUtil.FixParam(sEmail, True)
    End Set
  End Property

  Private iQBActivityID As Int32
  Private sInsUpdQBActivityID As String
  Property QBActivityID__Integer() As Int32
    Get
      Return iQBActivityID
    End Get
    Set(ByVal Value As Int32)
      iQBActivityID = Value
      sInsUpdQBActivityID = oUtil.FixParam(iQBActivityID, True)
    End Set
  End Property

  Private sAuthorizeNetURL As String
  Private sInsUpdAuthorizeNetURL As String
  Property AuthorizeNetURL__String() As String
    Get
      Return sAuthorizeNetURL
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetURL = Value
      sInsUpdAuthorizeNetURL = oUtil.FixParam(sAuthorizeNetURL, True)
    End Set
  End Property

  Private sAuthNETPaymentProfileID As String
  Private sInsUpdAuthNETPaymentProfileID As String
  Property AuthNETPaymentProfileID__String() As String
    Get
      Return sAuthNETPaymentProfileID
    End Get
    Set(ByVal Value As String)
      sAuthNETPaymentProfileID = Value
      sInsUpdAuthNETPaymentProfileID = oUtil.FixParam(sAuthNETPaymentProfileID, True)
    End Set
  End Property

  Public Sub Clear()
    iPaymentID = 0
    sInsUpdPaymentID = ""
    iBookingID = 0
    sInsUpdBookingID = ""
    iPropertyID = 0
    sInsUpdPropertyID = ""
    iHostID = 0
    sInsUpdHostID = ""
    iGuestID = 0
    sInsUpdGuestID = ""
    sCategory = ""
    sInsUpdCategory = ""
    sType = ""
    sInsUpdType = ""
    dAmount = 0.0
    sInsUpdAmount = ""
    sDate = ""
    sInsUpdDate = ""
    sComputer = ""
    sInsUpdComputer = ""
    sCheckNumber = ""
    sInsUpdCheckNumber = ""
    sCheckName = ""
    sInsUpdCheckName = ""
    sCCType = ""
    sInsUpdCCType = ""
    sCCNumber = ""
    sInsUpdCCNumber = ""
    sCCNumberEncrypted = ""
    sInsUpdCCNumberEncrypted = ""
    sCCName = ""
    sInsUpdCCName = ""
    sCCExpMonth = ""
    sInsUpdCCExpMonth = ""
    sCCExpYear = ""
    sInsUpdCCExpYear = ""
    sCCVerification = ""
    sInsUpdCCVerification = ""
    sCCConfirm = ""
    sInsUpdCCConfirm = ""
    sQBStatus = ""
    sInsUpdQBStatus = ""
    sNotes = ""
    sInsUpdNotes = ""
    sUser = ""
    sInsUpdUser = ""
    sStatus = ""
    sInsUpdStatus = ""
    iUseOnReceipt = 0
    sInsUpdUseOnReceipt = ""
    sOrigCCNumber = ""
    sInsUpdOrigCCNumber = ""
    sCCAddress = ""
    sInsUpdCCAddress = ""
    sCCCity = ""
    sInsUpdCCCity = ""
    sCCState = ""
    sInsUpdCCState = ""
    sCCZip = ""
    sInsUpdCCZip = ""
    sCCCountry = ""
    sInsUpdCCCountry = ""
    sAuthorizeNetReturnCode = ""
    sInsUpdAuthorizeNetReturnCode = ""
    sAuthorizeNetReasonCode = ""
    sInsUpdAuthorizeNetReasonCode = ""
    sAuthorizeNetReasonText = ""
    sInsUpdAuthorizeNetReasonText = ""
    sAuthorizeNetApprovalCode = ""
    sInsUpdAuthorizeNetApprovalCode = ""
    sAuthorizeNetAVSResultCode = ""
    sInsUpdAuthorizeNetAVSResultCode = ""
    sAuthorizeNetAVSResultText = ""
    sInsUpdAuthorizeNetAVSResultText = ""
    sAuthorizeNetTransactionID = ""
    sInsUpdAuthorizeNetTransactionID = ""
    sAuthorizeNetCVCCResponseCode = ""
    sInsUpdAuthorizeNetCVCCResponseCode = ""
    sAuthorizeNetCVCCResponseText = ""
    sInsUpdAuthorizeNetCVCCResponseText = ""
    sLocation = ""
    sInsUpdLocation = ""
    sBankName = ""
    sInsUpdBankName = ""
    sBankAccountType = ""
    sInsUpdBankAccountType = ""
    sBankAccountName = ""
    sInsUpdBankAccountName = ""
    sBankAccountNumber = ""
    sInsUpdBankAccountNumber = ""
    sBankAccountNumberEncrypted = ""
    sInsUpdBankAccountNumberEncrypted = ""
    sBankRoutingNumber = ""
    sInsUpdBankRoutingNumber = ""
    sEmail = ""
    sInsUpdEmail = ""
    iQBActivityID = 0
    sInsUpdQBActivityID = ""
    sAuthorizeNetURL = ""
    sInsUpdAuthorizeNetURL = ""
    sAuthNETPaymentProfileID = ""
    sInsUpdAuthNETPaymentProfileID = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdPaymentID.ToString = "'-12345'" Then sb.Append("PaymentID,")
        If sInsUpdBookingID.ToString = "'-12345'" Then sb.Append("BookingID,")
        If sInsUpdPropertyID.ToString = "'-12345'" Then sb.Append("PropertyID,")
        If sInsUpdHostID.ToString = "'-12345'" Then sb.Append("HostID,")
        If sInsUpdGuestID.ToString = "'-12345'" Then sb.Append("GuestID,")
        If sInsUpdCategory.ToString = "'Select'" Then sb.Append("Category,")
        If sInsUpdType.ToString = "'Select'" Then sb.Append("Type,")
        If sInsUpdAmount.ToString = "'-12345'" Then sb.Append("Amount,")
        If sInsUpdDate.ToString = "'Select'" Then sb.Append("[Date],")
        If sInsUpdComputer.ToString = "'Select'" Then sb.Append("Computer,")
        If sInsUpdCheckNumber.ToString = "'Select'" Then sb.Append("CheckNumber,")
        If sInsUpdCheckName.ToString = "'Select'" Then sb.Append("CheckName,")
        If sInsUpdCCType.ToString = "'Select'" Then sb.Append("CCType,")
        If sInsUpdCCNumber.ToString = "'Select'" Then sb.Append("CCNumber,")
        If sInsUpdCCNumberEncrypted.ToString = "'Select'" Then sb.Append("CCNumberEncrypted,")
        If sInsUpdCCName.ToString = "'Select'" Then sb.Append("CCName,")
        If sInsUpdCCExpMonth.ToString = "'Select'" Then sb.Append("CCExpMonth,")
        If sInsUpdCCExpYear.ToString = "'Select'" Then sb.Append("CCExpYear,")
        If sInsUpdCCVerification.ToString = "'Select'" Then sb.Append("CCVerification,")
        If sInsUpdCCConfirm.ToString = "'Select'" Then sb.Append("CCConfirm,")
        If sInsUpdQBStatus.ToString = "'Select'" Then sb.Append("QBStatus,")
        If sInsUpdNotes.ToString = "'Select'" Then sb.Append("Notes,")
        If sInsUpdUser.ToString = "'Select'" Then sb.Append("[User],")
        If sInsUpdStatus.ToString = "'Select'" Then sb.Append("Status,")
        If sInsUpdUseOnReceipt.ToString = "'-12345'" Then sb.Append("UseOnReceipt,")
        If sInsUpdOrigCCNumber.ToString = "'Select'" Then sb.Append("OrigCCNumber,")
        If sInsUpdCCAddress.ToString = "'Select'" Then sb.Append("CCAddress,")
        If sInsUpdCCCity.ToString = "'Select'" Then sb.Append("CCCity,")
        If sInsUpdCCState.ToString = "'Select'" Then sb.Append("CCState,")
        If sInsUpdCCZip.ToString = "'Select'" Then sb.Append("CCZip,")
        If sInsUpdCCCountry.ToString = "'Select'" Then sb.Append("CCCountry,")
        If sInsUpdAuthorizeNetReturnCode.ToString = "'Select'" Then sb.Append("AuthorizeNetReturnCode,")
        If sInsUpdAuthorizeNetReasonCode.ToString = "'Select'" Then sb.Append("AuthorizeNetReasonCode,")
        If sInsUpdAuthorizeNetReasonText.ToString = "'Select'" Then sb.Append("AuthorizeNetReasonText,")
        If sInsUpdAuthorizeNetApprovalCode.ToString = "'Select'" Then sb.Append("AuthorizeNetApprovalCode,")
        If sInsUpdAuthorizeNetAVSResultCode.ToString = "'Select'" Then sb.Append("AuthorizeNetAVSResultCode,")
        If sInsUpdAuthorizeNetAVSResultText.ToString = "'Select'" Then sb.Append("AuthorizeNetAVSResultText,")
        If sInsUpdAuthorizeNetTransactionID.ToString = "'Select'" Then sb.Append("AuthorizeNetTransactionID,")
        If sInsUpdAuthorizeNetCVCCResponseCode.ToString = "'Select'" Then sb.Append("AuthorizeNetCVCCResponseCode,")
        If sInsUpdAuthorizeNetCVCCResponseText.ToString = "'Select'" Then sb.Append("AuthorizeNetCVCCResponseText,")
        If sInsUpdLocation.ToString = "'Select'" Then sb.Append("Location,")
        If sInsUpdBankName.ToString = "'Select'" Then sb.Append("BankName,")
        If sInsUpdBankAccountType.ToString = "'Select'" Then sb.Append("BankAccountType,")
        If sInsUpdBankAccountName.ToString = "'Select'" Then sb.Append("BankAccountName,")
        If sInsUpdBankAccountNumber.ToString = "'Select'" Then sb.Append("BankAccountNumber,")
        If sInsUpdBankAccountNumberEncrypted.ToString = "'Select'" Then sb.Append("BankAccountNumberEncrypted,")
        If sInsUpdBankRoutingNumber.ToString = "'Select'" Then sb.Append("BankRoutingNumber,")
        If sInsUpdEmail.ToString = "'Select'" Then sb.Append("Email,")
        If sInsUpdQBActivityID.ToString = "'-12345'" Then sb.Append("QBActivityID,")
        If sInsUpdAuthorizeNetURL.ToString = "'Select'" Then sb.Append("AuthorizeNetURL,")
        If sInsUpdAuthNETPaymentProfileID.ToString = "'Select'" Then sb.Append("AuthNETPaymentProfileID,")
      Else
        sb.Append("PaymentID,")
        sb.Append("BookingID,")
        sb.Append("PropertyID,")
        sb.Append("HostID,")
        sb.Append("GuestID,")
        sb.Append("Category,")
        sb.Append("Type,")
        sb.Append("Amount,")
        sb.Append("[Date],")
        sb.Append("Computer,")
        sb.Append("CheckNumber,")
        sb.Append("CheckName,")
        sb.Append("CCType,")
        sb.Append("CCNumber,")
        sb.Append("CCNumberEncrypted,")
        sb.Append("CCName,")
        sb.Append("CCExpMonth,")
        sb.Append("CCExpYear,")
        sb.Append("CCVerification,")
        sb.Append("CCConfirm,")
        sb.Append("QBStatus,")
        sb.Append("Notes,")
        sb.Append("[User],")
        sb.Append("Status,")
        sb.Append("UseOnReceipt,")
        sb.Append("OrigCCNumber,")
        sb.Append("CCAddress,")
        sb.Append("CCCity,")
        sb.Append("CCState,")
        sb.Append("CCZip,")
        sb.Append("CCCountry,")
        sb.Append("AuthorizeNetReturnCode,")
        sb.Append("AuthorizeNetReasonCode,")
        sb.Append("AuthorizeNetReasonText,")
        sb.Append("AuthorizeNetApprovalCode,")
        sb.Append("AuthorizeNetAVSResultCode,")
        sb.Append("AuthorizeNetAVSResultText,")
        sb.Append("AuthorizeNetTransactionID,")
        sb.Append("AuthorizeNetCVCCResponseCode,")
        sb.Append("AuthorizeNetCVCCResponseText,")
        sb.Append("Location,")
        sb.Append("BankName,")
        sb.Append("BankAccountType,")
        sb.Append("BankAccountName,")
        sb.Append("BankAccountNumber,")
        sb.Append("BankAccountNumberEncrypted,")
        sb.Append("BankRoutingNumber,")
        sb.Append("Email,")
        sb.Append("QBActivityID,")
        sb.Append("AuthorizeNetURL,")
        sb.Append("AuthNETPaymentProfileID,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [DepositPayments]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdPaymentID.ToString <> "") And (sInsUpdPaymentID <> "'-12345'") Then sbw.Append("PaymentID=" & sInsUpdPaymentID & " and ")
      If (sInsUpdBookingID.ToString <> "") And (sInsUpdBookingID <> "'-12345'") Then sbw.Append("BookingID=" & sInsUpdBookingID & " and ")
      If (sInsUpdPropertyID.ToString <> "") And (sInsUpdPropertyID <> "'-12345'") Then sbw.Append("PropertyID=" & sInsUpdPropertyID & " and ")
      If (sInsUpdHostID.ToString <> "") And (sInsUpdHostID <> "'-12345'") Then sbw.Append("HostID=" & sInsUpdHostID & " and ")
      If (sInsUpdGuestID.ToString <> "") And (sInsUpdGuestID <> "'-12345'") Then sbw.Append("GuestID=" & sInsUpdGuestID & " and ")
      If (sInsUpdCategory.ToString <> "") And (sInsUpdCategory <> "'Select'") Then sbw.Append("Category=" & sInsUpdCategory & " and ")
      If (sInsUpdType.ToString <> "") And (sInsUpdType <> "'Select'") Then sbw.Append("Type=" & sInsUpdType & " and ")
      If (sInsUpdAmount.ToString <> "") And (sInsUpdAmount <> "'-12345'") Then sbw.Append("Amount=" & sInsUpdAmount & " and ")
      If (sInsUpdDate.ToString <> "") And (sInsUpdDate <> "'Select'") Then sbw.Append("[Date]=" & sInsUpdDate & " and ")
      If (sInsUpdComputer.ToString <> "") And (sInsUpdComputer <> "'Select'") Then sbw.Append("Computer=" & sInsUpdComputer & " and ")
      If (sInsUpdCheckNumber.ToString <> "") And (sInsUpdCheckNumber <> "'Select'") Then sbw.Append("CheckNumber=" & sInsUpdCheckNumber & " and ")
      If (sInsUpdCheckName.ToString <> "") And (sInsUpdCheckName <> "'Select'") Then sbw.Append("CheckName=" & sInsUpdCheckName & " and ")
      If (sInsUpdCCType.ToString <> "") And (sInsUpdCCType <> "'Select'") Then sbw.Append("CCType=" & sInsUpdCCType & " and ")
      If (sInsUpdCCNumber.ToString <> "") And (sInsUpdCCNumber <> "'Select'") Then sbw.Append("CCNumber=" & sInsUpdCCNumber & " and ")
      If (sInsUpdCCNumberEncrypted.ToString <> "") And (sInsUpdCCNumberEncrypted <> "'Select'") Then sbw.Append("CCNumberEncrypted=" & sInsUpdCCNumberEncrypted & " and ")
      If (sInsUpdCCName.ToString <> "") And (sInsUpdCCName <> "'Select'") Then sbw.Append("CCName=" & sInsUpdCCName & " and ")
      If (sInsUpdCCExpMonth.ToString <> "") And (sInsUpdCCExpMonth <> "'Select'") Then sbw.Append("CCExpMonth=" & sInsUpdCCExpMonth & " and ")
      If (sInsUpdCCExpYear.ToString <> "") And (sInsUpdCCExpYear <> "'Select'") Then sbw.Append("CCExpYear=" & sInsUpdCCExpYear & " and ")
      If (sInsUpdCCVerification.ToString <> "") And (sInsUpdCCVerification <> "'Select'") Then sbw.Append("CCVerification=" & sInsUpdCCVerification & " and ")
      If (sInsUpdCCConfirm.ToString <> "") And (sInsUpdCCConfirm <> "'Select'") Then sbw.Append("CCConfirm=" & sInsUpdCCConfirm & " and ")
      If (sInsUpdQBStatus.ToString <> "") And (sInsUpdQBStatus <> "'Select'") Then sbw.Append("QBStatus=" & sInsUpdQBStatus & " and ")
      If (sInsUpdNotes.ToString <> "") And (sInsUpdNotes <> "'Select'") Then sbw.Append("Notes=" & sInsUpdNotes & " and ")
      If (sInsUpdUser.ToString <> "") And (sInsUpdUser <> "'Select'") Then sbw.Append("[User]=" & sInsUpdUser & " and ")
      If (sInsUpdStatus.ToString <> "") And (sInsUpdStatus <> "'Select'") Then sbw.Append("Status=" & sInsUpdStatus & " and ")
      If (sInsUpdUseOnReceipt.ToString <> "") And (sInsUpdUseOnReceipt <> "'-12345'") Then sbw.Append("UseOnReceipt=" & sInsUpdUseOnReceipt & " and ")
      If (sInsUpdOrigCCNumber.ToString <> "") And (sInsUpdOrigCCNumber <> "'Select'") Then sbw.Append("OrigCCNumber=" & sInsUpdOrigCCNumber & " and ")
      If (sInsUpdCCAddress.ToString <> "") And (sInsUpdCCAddress <> "'Select'") Then sbw.Append("CCAddress=" & sInsUpdCCAddress & " and ")
      If (sInsUpdCCCity.ToString <> "") And (sInsUpdCCCity <> "'Select'") Then sbw.Append("CCCity=" & sInsUpdCCCity & " and ")
      If (sInsUpdCCState.ToString <> "") And (sInsUpdCCState <> "'Select'") Then sbw.Append("CCState=" & sInsUpdCCState & " and ")
      If (sInsUpdCCZip.ToString <> "") And (sInsUpdCCZip <> "'Select'") Then sbw.Append("CCZip=" & sInsUpdCCZip & " and ")
      If (sInsUpdCCCountry.ToString <> "") And (sInsUpdCCCountry <> "'Select'") Then sbw.Append("CCCountry=" & sInsUpdCCCountry & " and ")
      If (sInsUpdAuthorizeNetReturnCode.ToString <> "") And (sInsUpdAuthorizeNetReturnCode <> "'Select'") Then sbw.Append("AuthorizeNetReturnCode=" & sInsUpdAuthorizeNetReturnCode & " and ")
      If (sInsUpdAuthorizeNetReasonCode.ToString <> "") And (sInsUpdAuthorizeNetReasonCode <> "'Select'") Then sbw.Append("AuthorizeNetReasonCode=" & sInsUpdAuthorizeNetReasonCode & " and ")
      If (sInsUpdAuthorizeNetReasonText.ToString <> "") And (sInsUpdAuthorizeNetReasonText <> "'Select'") Then sbw.Append("AuthorizeNetReasonText=" & sInsUpdAuthorizeNetReasonText & " and ")
      If (sInsUpdAuthorizeNetApprovalCode.ToString <> "") And (sInsUpdAuthorizeNetApprovalCode <> "'Select'") Then sbw.Append("AuthorizeNetApprovalCode=" & sInsUpdAuthorizeNetApprovalCode & " and ")
      If (sInsUpdAuthorizeNetAVSResultCode.ToString <> "") And (sInsUpdAuthorizeNetAVSResultCode <> "'Select'") Then sbw.Append("AuthorizeNetAVSResultCode=" & sInsUpdAuthorizeNetAVSResultCode & " and ")
      If (sInsUpdAuthorizeNetAVSResultText.ToString <> "") And (sInsUpdAuthorizeNetAVSResultText <> "'Select'") Then sbw.Append("AuthorizeNetAVSResultText=" & sInsUpdAuthorizeNetAVSResultText & " and ")
      If (sInsUpdAuthorizeNetTransactionID.ToString <> "") And (sInsUpdAuthorizeNetTransactionID <> "'Select'") Then sbw.Append("AuthorizeNetTransactionID=" & sInsUpdAuthorizeNetTransactionID & " and ")
      If (sInsUpdAuthorizeNetCVCCResponseCode.ToString <> "") And (sInsUpdAuthorizeNetCVCCResponseCode <> "'Select'") Then sbw.Append("AuthorizeNetCVCCResponseCode=" & sInsUpdAuthorizeNetCVCCResponseCode & " and ")
      If (sInsUpdAuthorizeNetCVCCResponseText.ToString <> "") And (sInsUpdAuthorizeNetCVCCResponseText <> "'Select'") Then sbw.Append("AuthorizeNetCVCCResponseText=" & sInsUpdAuthorizeNetCVCCResponseText & " and ")
      If (sInsUpdLocation.ToString <> "") And (sInsUpdLocation <> "'Select'") Then sbw.Append("Location=" & sInsUpdLocation & " and ")
      If (sInsUpdBankName.ToString <> "") And (sInsUpdBankName <> "'Select'") Then sbw.Append("BankName=" & sInsUpdBankName & " and ")
      If (sInsUpdBankAccountType.ToString <> "") And (sInsUpdBankAccountType <> "'Select'") Then sbw.Append("BankAccountType=" & sInsUpdBankAccountType & " and ")
      If (sInsUpdBankAccountName.ToString <> "") And (sInsUpdBankAccountName <> "'Select'") Then sbw.Append("BankAccountName=" & sInsUpdBankAccountName & " and ")
      If (sInsUpdBankAccountNumber.ToString <> "") And (sInsUpdBankAccountNumber <> "'Select'") Then sbw.Append("BankAccountNumber=" & sInsUpdBankAccountNumber & " and ")
      If (sInsUpdBankAccountNumberEncrypted.ToString <> "") And (sInsUpdBankAccountNumberEncrypted <> "'Select'") Then sbw.Append("BankAccountNumberEncrypted=" & sInsUpdBankAccountNumberEncrypted & " and ")
      If (sInsUpdBankRoutingNumber.ToString <> "") And (sInsUpdBankRoutingNumber <> "'Select'") Then sbw.Append("BankRoutingNumber=" & sInsUpdBankRoutingNumber & " and ")
      If (sInsUpdEmail.ToString <> "") And (sInsUpdEmail <> "'Select'") Then sbw.Append("Email=" & sInsUpdEmail & " and ")
      If (sInsUpdQBActivityID.ToString <> "") And (sInsUpdQBActivityID <> "'-12345'") Then sbw.Append("QBActivityID=" & sInsUpdQBActivityID & " and ")
      If (sInsUpdAuthorizeNetURL.ToString <> "") And (sInsUpdAuthorizeNetURL <> "'Select'") Then sbw.Append("AuthorizeNetURL=" & sInsUpdAuthorizeNetURL & " and ")
      If (sInsUpdAuthNETPaymentProfileID.ToString <> "") And (sInsUpdAuthNETPaymentProfileID <> "'Select'") Then sbw.Append("AuthNETPaymentProfileID=" & sInsUpdAuthNETPaymentProfileID & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      PaymentID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("PaymentID")), 0, CurrentRow.Item("PaymentID").ToString)
      BookingID__Integer = IIf(IsDBNull(CurrentRow.Item("BookingID")), 0, CurrentRow.Item("BookingID"))
      PropertyID__Integer = IIf(IsDBNull(CurrentRow.Item("PropertyID")), 0, CurrentRow.Item("PropertyID"))
      HostID__Integer = IIf(IsDBNull(CurrentRow.Item("HostID")), 0, CurrentRow.Item("HostID"))
      GuestID__Integer = IIf(IsDBNull(CurrentRow.Item("GuestID")), 0, CurrentRow.Item("GuestID"))
      Category__String = IIf(IsDBNull(CurrentRow.Item("Category")), "", CurrentRow.Item("Category"))
      Type__String = IIf(IsDBNull(CurrentRow.Item("Type")), "", CurrentRow.Item("Type"))
      Amount__Numeric = IIf(IsDBNull(CurrentRow.Item("Amount")), 0.0, CurrentRow.Item("Amount"))
      Date__Date = IIf(IsDBNull(CurrentRow.Item("Date")), "", CurrentRow.Item("Date"))
      Computer__String = IIf(IsDBNull(CurrentRow.Item("Computer")), "", CurrentRow.Item("Computer"))
      CheckNumber__String = IIf(IsDBNull(CurrentRow.Item("CheckNumber")), "", CurrentRow.Item("CheckNumber"))
      CheckName__String = IIf(IsDBNull(CurrentRow.Item("CheckName")), "", CurrentRow.Item("CheckName"))
      CCType__String = IIf(IsDBNull(CurrentRow.Item("CCType")), "", CurrentRow.Item("CCType"))
      CCNumber__String = IIf(IsDBNull(CurrentRow.Item("CCNumber")), "", CurrentRow.Item("CCNumber"))
      CCNumberEncrypted__String = IIf(IsDBNull(CurrentRow.Item("CCNumberEncrypted")), "", CurrentRow.Item("CCNumberEncrypted"))
      CCName__String = IIf(IsDBNull(CurrentRow.Item("CCName")), "", CurrentRow.Item("CCName"))
      CCExpMonth__String = IIf(IsDBNull(CurrentRow.Item("CCExpMonth")), "", CurrentRow.Item("CCExpMonth"))
      CCExpYear__String = IIf(IsDBNull(CurrentRow.Item("CCExpYear")), "", CurrentRow.Item("CCExpYear"))
      CCVerification__String = IIf(IsDBNull(CurrentRow.Item("CCVerification")), "", CurrentRow.Item("CCVerification"))
      CCConfirm__String = IIf(IsDBNull(CurrentRow.Item("CCConfirm")), "", CurrentRow.Item("CCConfirm"))
      QBStatus__String = IIf(IsDBNull(CurrentRow.Item("QBStatus")), "", CurrentRow.Item("QBStatus"))
      Notes__String = IIf(IsDBNull(CurrentRow.Item("Notes")), "", CurrentRow.Item("Notes"))
      User__String = IIf(IsDBNull(CurrentRow.Item("User")), "", CurrentRow.Item("User"))
      Status__String = IIf(IsDBNull(CurrentRow.Item("Status")), "", CurrentRow.Item("Status"))
      UseOnReceipt__Integer = IIf(IsDBNull(CurrentRow.Item("UseOnReceipt")), 0, CurrentRow.Item("UseOnReceipt"))
      OrigCCNumber__String = IIf(IsDBNull(CurrentRow.Item("OrigCCNumber")), "", CurrentRow.Item("OrigCCNumber"))
      CCAddress__String = IIf(IsDBNull(CurrentRow.Item("CCAddress")), "", CurrentRow.Item("CCAddress"))
      CCCity__String = IIf(IsDBNull(CurrentRow.Item("CCCity")), "", CurrentRow.Item("CCCity"))
      CCState__String = IIf(IsDBNull(CurrentRow.Item("CCState")), "", CurrentRow.Item("CCState"))
      CCZip__String = IIf(IsDBNull(CurrentRow.Item("CCZip")), "", CurrentRow.Item("CCZip"))
      CCCountry__String = IIf(IsDBNull(CurrentRow.Item("CCCountry")), "", CurrentRow.Item("CCCountry"))
      AuthorizeNetReturnCode__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetReturnCode")), "", CurrentRow.Item("AuthorizeNetReturnCode"))
      AuthorizeNetReasonCode__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetReasonCode")), "", CurrentRow.Item("AuthorizeNetReasonCode"))
      AuthorizeNetReasonText__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetReasonText")), "", CurrentRow.Item("AuthorizeNetReasonText"))
      AuthorizeNetApprovalCode__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetApprovalCode")), "", CurrentRow.Item("AuthorizeNetApprovalCode"))
      AuthorizeNetAVSResultCode__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetAVSResultCode")), "", CurrentRow.Item("AuthorizeNetAVSResultCode"))
      AuthorizeNetAVSResultText__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetAVSResultText")), "", CurrentRow.Item("AuthorizeNetAVSResultText"))
      AuthorizeNetTransactionID__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetTransactionID")), "", CurrentRow.Item("AuthorizeNetTransactionID"))
      AuthorizeNetCVCCResponseCode__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetCVCCResponseCode")), "", CurrentRow.Item("AuthorizeNetCVCCResponseCode"))
      AuthorizeNetCVCCResponseText__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetCVCCResponseText")), "", CurrentRow.Item("AuthorizeNetCVCCResponseText"))
      Location__String = IIf(IsDBNull(CurrentRow.Item("Location")), "", CurrentRow.Item("Location"))
      BankName__String = IIf(IsDBNull(CurrentRow.Item("BankName")), "", CurrentRow.Item("BankName"))
      BankAccountType__String = IIf(IsDBNull(CurrentRow.Item("BankAccountType")), "", CurrentRow.Item("BankAccountType"))
      BankAccountName__String = IIf(IsDBNull(CurrentRow.Item("BankAccountName")), "", CurrentRow.Item("BankAccountName"))
      BankAccountNumber__String = IIf(IsDBNull(CurrentRow.Item("BankAccountNumber")), "", CurrentRow.Item("BankAccountNumber"))
      BankAccountNumberEncrypted__String = IIf(IsDBNull(CurrentRow.Item("BankAccountNumberEncrypted")), "", CurrentRow.Item("BankAccountNumberEncrypted"))
      BankRoutingNumber__String = IIf(IsDBNull(CurrentRow.Item("BankRoutingNumber")), "", CurrentRow.Item("BankRoutingNumber"))
      Email__String = IIf(IsDBNull(CurrentRow.Item("Email")), "", CurrentRow.Item("Email"))
      QBActivityID__Integer = IIf(IsDBNull(CurrentRow.Item("QBActivityID")), 0, CurrentRow.Item("QBActivityID"))
      AuthorizeNetURL__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetURL")), "", CurrentRow.Item("AuthorizeNetURL"))
      AuthNETPaymentProfileID__String = IIf(IsDBNull(CurrentRow.Item("AuthNETPaymentProfileID")), "", CurrentRow.Item("AuthNETPaymentProfileID"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [DepositPayments](")
    If sInsUpdBookingID.ToString <> "" Then
      sb.Append("BookingID,")
      sbv.Append(sInsUpdBookingID & ",")
    End If
    If sInsUpdPropertyID.ToString <> "" Then
      sb.Append("PropertyID,")
      sbv.Append(sInsUpdPropertyID & ",")
    End If
    If sInsUpdHostID.ToString <> "" Then
      sb.Append("HostID,")
      sbv.Append(sInsUpdHostID & ",")
    End If
    If sInsUpdGuestID.ToString <> "" Then
      sb.Append("GuestID,")
      sbv.Append(sInsUpdGuestID & ",")
    End If
    If sInsUpdCategory.ToString <> "" Then
      sb.Append("Category,")
      sbv.Append(sInsUpdCategory & ",")
    End If
    If sInsUpdType.ToString <> "" Then
      sb.Append("Type,")
      sbv.Append(sInsUpdType & ",")
    End If
    If sInsUpdAmount.ToString <> "" Then
      sb.Append("Amount,")
      sbv.Append(sInsUpdAmount & ",")
    End If
    If sInsUpdDate.ToString <> "" Then
      sb.Append("[Date],")
      sbv.Append(sInsUpdDate & ",")
    End If
    If sInsUpdComputer.ToString <> "" Then
      sb.Append("Computer,")
      sbv.Append(sInsUpdComputer & ",")
    End If
    If sInsUpdCheckNumber.ToString <> "" Then
      sb.Append("CheckNumber,")
      sbv.Append(sInsUpdCheckNumber & ",")
    End If
    If sInsUpdCheckName.ToString <> "" Then
      sb.Append("CheckName,")
      sbv.Append(sInsUpdCheckName & ",")
    End If
    If sInsUpdCCType.ToString <> "" Then
      sb.Append("CCType,")
      sbv.Append(sInsUpdCCType & ",")
    End If
    If sInsUpdCCNumber.ToString <> "" Then
      sb.Append("CCNumber,")
      sbv.Append(sInsUpdCCNumber & ",")
    End If
    If sInsUpdCCNumberEncrypted.ToString <> "" Then
      sb.Append("CCNumberEncrypted,")
      sbv.Append(sInsUpdCCNumberEncrypted & ",")
    End If
    If sInsUpdCCName.ToString <> "" Then
      sb.Append("CCName,")
      sbv.Append(sInsUpdCCName & ",")
    End If
    If sInsUpdCCExpMonth.ToString <> "" Then
      sb.Append("CCExpMonth,")
      sbv.Append(sInsUpdCCExpMonth & ",")
    End If
    If sInsUpdCCExpYear.ToString <> "" Then
      sb.Append("CCExpYear,")
      sbv.Append(sInsUpdCCExpYear & ",")
    End If
    If sInsUpdCCVerification.ToString <> "" Then
      sb.Append("CCVerification,")
      sbv.Append(sInsUpdCCVerification & ",")
    End If
    If sInsUpdCCConfirm.ToString <> "" Then
      sb.Append("CCConfirm,")
      sbv.Append(sInsUpdCCConfirm & ",")
    End If
    If sInsUpdQBStatus.ToString <> "" Then
      sb.Append("QBStatus,")
      sbv.Append(sInsUpdQBStatus & ",")
    End If
    If sInsUpdNotes.ToString <> "" Then
      sb.Append("Notes,")
      sbv.Append(sInsUpdNotes & ",")
    End If
    If sInsUpdUser.ToString <> "" Then
      sb.Append("[User],")
      sbv.Append(sInsUpdUser & ",")
    End If
    If sInsUpdStatus.ToString <> "" Then
      sb.Append("Status,")
      sbv.Append(sInsUpdStatus & ",")
    End If
    If sInsUpdUseOnReceipt.ToString <> "" Then
      sb.Append("UseOnReceipt,")
      sbv.Append(sInsUpdUseOnReceipt & ",")
    End If
    If sInsUpdOrigCCNumber.ToString <> "" Then
      sb.Append("OrigCCNumber,")
      sbv.Append(sInsUpdOrigCCNumber & ",")
    End If
    If sInsUpdCCAddress.ToString <> "" Then
      sb.Append("CCAddress,")
      sbv.Append(sInsUpdCCAddress & ",")
    End If
    If sInsUpdCCCity.ToString <> "" Then
      sb.Append("CCCity,")
      sbv.Append(sInsUpdCCCity & ",")
    End If
    If sInsUpdCCState.ToString <> "" Then
      sb.Append("CCState,")
      sbv.Append(sInsUpdCCState & ",")
    End If
    If sInsUpdCCZip.ToString <> "" Then
      sb.Append("CCZip,")
      sbv.Append(sInsUpdCCZip & ",")
    End If
    If sInsUpdCCCountry.ToString <> "" Then
      sb.Append("CCCountry,")
      sbv.Append(sInsUpdCCCountry & ",")
    End If
    If sInsUpdAuthorizeNetReturnCode.ToString <> "" Then
      sb.Append("AuthorizeNetReturnCode,")
      sbv.Append(sInsUpdAuthorizeNetReturnCode & ",")
    End If
    If sInsUpdAuthorizeNetReasonCode.ToString <> "" Then
      sb.Append("AuthorizeNetReasonCode,")
      sbv.Append(sInsUpdAuthorizeNetReasonCode & ",")
    End If
    If sInsUpdAuthorizeNetReasonText.ToString <> "" Then
      sb.Append("AuthorizeNetReasonText,")
      sbv.Append(sInsUpdAuthorizeNetReasonText & ",")
    End If
    If sInsUpdAuthorizeNetApprovalCode.ToString <> "" Then
      sb.Append("AuthorizeNetApprovalCode,")
      sbv.Append(sInsUpdAuthorizeNetApprovalCode & ",")
    End If
    If sInsUpdAuthorizeNetAVSResultCode.ToString <> "" Then
      sb.Append("AuthorizeNetAVSResultCode,")
      sbv.Append(sInsUpdAuthorizeNetAVSResultCode & ",")
    End If
    If sInsUpdAuthorizeNetAVSResultText.ToString <> "" Then
      sb.Append("AuthorizeNetAVSResultText,")
      sbv.Append(sInsUpdAuthorizeNetAVSResultText & ",")
    End If
    If sInsUpdAuthorizeNetTransactionID.ToString <> "" Then
      sb.Append("AuthorizeNetTransactionID,")
      sbv.Append(sInsUpdAuthorizeNetTransactionID & ",")
    End If
    If sInsUpdAuthorizeNetCVCCResponseCode.ToString <> "" Then
      sb.Append("AuthorizeNetCVCCResponseCode,")
      sbv.Append(sInsUpdAuthorizeNetCVCCResponseCode & ",")
    End If
    If sInsUpdAuthorizeNetCVCCResponseText.ToString <> "" Then
      sb.Append("AuthorizeNetCVCCResponseText,")
      sbv.Append(sInsUpdAuthorizeNetCVCCResponseText & ",")
    End If
    If sInsUpdLocation.ToString <> "" Then
      sb.Append("Location,")
      sbv.Append(sInsUpdLocation & ",")
    End If
    If sInsUpdBankName.ToString <> "" Then
      sb.Append("BankName,")
      sbv.Append(sInsUpdBankName & ",")
    End If
    If sInsUpdBankAccountType.ToString <> "" Then
      sb.Append("BankAccountType,")
      sbv.Append(sInsUpdBankAccountType & ",")
    End If
    If sInsUpdBankAccountName.ToString <> "" Then
      sb.Append("BankAccountName,")
      sbv.Append(sInsUpdBankAccountName & ",")
    End If
    If sInsUpdBankAccountNumber.ToString <> "" Then
      sb.Append("BankAccountNumber,")
      sbv.Append(sInsUpdBankAccountNumber & ",")
    End If
    If sInsUpdBankAccountNumberEncrypted.ToString <> "" Then
      sb.Append("BankAccountNumberEncrypted,")
      sbv.Append(sInsUpdBankAccountNumberEncrypted & ",")
    End If
    If sInsUpdBankRoutingNumber.ToString <> "" Then
      sb.Append("BankRoutingNumber,")
      sbv.Append(sInsUpdBankRoutingNumber & ",")
    End If
    If sInsUpdEmail.ToString <> "" Then
      sb.Append("Email,")
      sbv.Append(sInsUpdEmail & ",")
    End If
    If sInsUpdQBActivityID.ToString <> "" Then
      sb.Append("QBActivityID,")
      sbv.Append(sInsUpdQBActivityID & ",")
    End If
    If sInsUpdAuthorizeNetURL.ToString <> "" Then
      sb.Append("AuthorizeNetURL,")
      sbv.Append(sInsUpdAuthorizeNetURL & ",")
    End If
    If sInsUpdAuthNETPaymentProfileID.ToString <> "" Then
      sb.Append("AuthNETPaymentProfileID,")
      sbv.Append(sInsUpdAuthNETPaymentProfileID & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(PaymentID) from [DepositPayments]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    PaymentID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [DepositPayments] Set ")
    If sInsUpdBookingID.ToString <> "" Then sb.Append("BookingID=" & sInsUpdBookingID & ",")
    If sInsUpdPropertyID.ToString <> "" Then sb.Append("PropertyID=" & sInsUpdPropertyID & ",")
    If sInsUpdHostID.ToString <> "" Then sb.Append("HostID=" & sInsUpdHostID & ",")
    If sInsUpdGuestID.ToString <> "" Then sb.Append("GuestID=" & sInsUpdGuestID & ",")
    If sInsUpdCategory.ToString <> "" Then sb.Append("Category=" & sInsUpdCategory & ",")
    If sInsUpdType.ToString <> "" Then sb.Append("Type=" & sInsUpdType & ",")
    If sInsUpdAmount.ToString <> "" Then sb.Append("Amount=" & sInsUpdAmount & ",")
    If sInsUpdDate.ToString <> "" Then sb.Append("[Date]=" & sInsUpdDate & ",")
    If sInsUpdComputer.ToString <> "" Then sb.Append("Computer=" & sInsUpdComputer & ",")
    If sInsUpdCheckNumber.ToString <> "" Then sb.Append("CheckNumber=" & sInsUpdCheckNumber & ",")
    If sInsUpdCheckName.ToString <> "" Then sb.Append("CheckName=" & sInsUpdCheckName & ",")
    If sInsUpdCCType.ToString <> "" Then sb.Append("CCType=" & sInsUpdCCType & ",")
    If sInsUpdCCNumber.ToString <> "" Then sb.Append("CCNumber=" & sInsUpdCCNumber & ",")
    If sInsUpdCCNumberEncrypted.ToString <> "" Then sb.Append("CCNumberEncrypted=" & sInsUpdCCNumberEncrypted & ",")
    If sInsUpdCCName.ToString <> "" Then sb.Append("CCName=" & sInsUpdCCName & ",")
    If sInsUpdCCExpMonth.ToString <> "" Then sb.Append("CCExpMonth=" & sInsUpdCCExpMonth & ",")
    If sInsUpdCCExpYear.ToString <> "" Then sb.Append("CCExpYear=" & sInsUpdCCExpYear & ",")
    If sInsUpdCCVerification.ToString <> "" Then sb.Append("CCVerification=" & sInsUpdCCVerification & ",")
    If sInsUpdCCConfirm.ToString <> "" Then sb.Append("CCConfirm=" & sInsUpdCCConfirm & ",")
    If sInsUpdQBStatus.ToString <> "" Then sb.Append("QBStatus=" & sInsUpdQBStatus & ",")
    If sInsUpdNotes.ToString <> "" Then sb.Append("Notes=" & sInsUpdNotes & ",")
    If sInsUpdUser.ToString <> "" Then sb.Append("[User]=" & sInsUpdUser & ",")
    If sInsUpdStatus.ToString <> "" Then sb.Append("Status=" & sInsUpdStatus & ",")
    If sInsUpdUseOnReceipt.ToString <> "" Then sb.Append("UseOnReceipt=" & sInsUpdUseOnReceipt & ",")
    If sInsUpdOrigCCNumber.ToString <> "" Then sb.Append("OrigCCNumber=" & sInsUpdOrigCCNumber & ",")
    If sInsUpdCCAddress.ToString <> "" Then sb.Append("CCAddress=" & sInsUpdCCAddress & ",")
    If sInsUpdCCCity.ToString <> "" Then sb.Append("CCCity=" & sInsUpdCCCity & ",")
    If sInsUpdCCState.ToString <> "" Then sb.Append("CCState=" & sInsUpdCCState & ",")
    If sInsUpdCCZip.ToString <> "" Then sb.Append("CCZip=" & sInsUpdCCZip & ",")
    If sInsUpdCCCountry.ToString <> "" Then sb.Append("CCCountry=" & sInsUpdCCCountry & ",")
    If sInsUpdAuthorizeNetReturnCode.ToString <> "" Then sb.Append("AuthorizeNetReturnCode=" & sInsUpdAuthorizeNetReturnCode & ",")
    If sInsUpdAuthorizeNetReasonCode.ToString <> "" Then sb.Append("AuthorizeNetReasonCode=" & sInsUpdAuthorizeNetReasonCode & ",")
    If sInsUpdAuthorizeNetReasonText.ToString <> "" Then sb.Append("AuthorizeNetReasonText=" & sInsUpdAuthorizeNetReasonText & ",")
    If sInsUpdAuthorizeNetApprovalCode.ToString <> "" Then sb.Append("AuthorizeNetApprovalCode=" & sInsUpdAuthorizeNetApprovalCode & ",")
    If sInsUpdAuthorizeNetAVSResultCode.ToString <> "" Then sb.Append("AuthorizeNetAVSResultCode=" & sInsUpdAuthorizeNetAVSResultCode & ",")
    If sInsUpdAuthorizeNetAVSResultText.ToString <> "" Then sb.Append("AuthorizeNetAVSResultText=" & sInsUpdAuthorizeNetAVSResultText & ",")
    If sInsUpdAuthorizeNetTransactionID.ToString <> "" Then sb.Append("AuthorizeNetTransactionID=" & sInsUpdAuthorizeNetTransactionID & ",")
    If sInsUpdAuthorizeNetCVCCResponseCode.ToString <> "" Then sb.Append("AuthorizeNetCVCCResponseCode=" & sInsUpdAuthorizeNetCVCCResponseCode & ",")
    If sInsUpdAuthorizeNetCVCCResponseText.ToString <> "" Then sb.Append("AuthorizeNetCVCCResponseText=" & sInsUpdAuthorizeNetCVCCResponseText & ",")
    If sInsUpdLocation.ToString <> "" Then sb.Append("Location=" & sInsUpdLocation & ",")
    If sInsUpdBankName.ToString <> "" Then sb.Append("BankName=" & sInsUpdBankName & ",")
    If sInsUpdBankAccountType.ToString <> "" Then sb.Append("BankAccountType=" & sInsUpdBankAccountType & ",")
    If sInsUpdBankAccountName.ToString <> "" Then sb.Append("BankAccountName=" & sInsUpdBankAccountName & ",")
    If sInsUpdBankAccountNumber.ToString <> "" Then sb.Append("BankAccountNumber=" & sInsUpdBankAccountNumber & ",")
    If sInsUpdBankAccountNumberEncrypted.ToString <> "" Then sb.Append("BankAccountNumberEncrypted=" & sInsUpdBankAccountNumberEncrypted & ",")
    If sInsUpdBankRoutingNumber.ToString <> "" Then sb.Append("BankRoutingNumber=" & sInsUpdBankRoutingNumber & ",")
    If sInsUpdEmail.ToString <> "" Then sb.Append("Email=" & sInsUpdEmail & ",")
    If sInsUpdQBActivityID.ToString <> "" Then sb.Append("QBActivityID=" & sInsUpdQBActivityID & ",")
    If sInsUpdAuthorizeNetURL.ToString <> "" Then sb.Append("AuthorizeNetURL=" & sInsUpdAuthorizeNetURL & ",")
    If sInsUpdAuthNETPaymentProfileID.ToString <> "" Then sb.Append("AuthNETPaymentProfileID=" & sInsUpdAuthNETPaymentProfileID & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where PaymentID=" & sInsUpdPaymentID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [DepositPayments] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("PaymentID=" & sInsUpdPaymentID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TableRefundPayments

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iPaymentID As Int32
  Private sInsUpdPaymentID As String
  Property PaymentID_PK__Integer() As Int32
    Get
      Return iPaymentID
    End Get
    Set(ByVal Value As Int32)
      iPaymentID = Value
      sInsUpdPaymentID = oUtil.FixParam(iPaymentID, True)
    End Set
  End Property

  Private iBookingID As Int32
  Private sInsUpdBookingID As String
  Property BookingID__Integer() As Int32
    Get
      Return iBookingID
    End Get
    Set(ByVal Value As Int32)
      iBookingID = Value
      sInsUpdBookingID = oUtil.FixParam(iBookingID, True)
    End Set
  End Property

  Private iPropertyID As Int32
  Private sInsUpdPropertyID As String
  Property PropertyID__Integer() As Int32
    Get
      Return iPropertyID
    End Get
    Set(ByVal Value As Int32)
      iPropertyID = Value
      sInsUpdPropertyID = oUtil.FixParam(iPropertyID, True)
    End Set
  End Property

  Private iHostID As Int32
  Private sInsUpdHostID As String
  Property HostID__Integer() As Int32
    Get
      Return iHostID
    End Get
    Set(ByVal Value As Int32)
      iHostID = Value
      sInsUpdHostID = oUtil.FixParam(iHostID, True)
    End Set
  End Property

  Private iGuestID As Int32
  Private sInsUpdGuestID As String
  Property GuestID__Integer() As Int32
    Get
      Return iGuestID
    End Get
    Set(ByVal Value As Int32)
      iGuestID = Value
      sInsUpdGuestID = oUtil.FixParam(iGuestID, True)
    End Set
  End Property

  Private sCategory As String
  Private sInsUpdCategory As String
  Property Category__String() As String
    Get
      Return sCategory
    End Get
    Set(ByVal Value As String)
      sCategory = Value
      sInsUpdCategory = oUtil.FixParam(sCategory, True)
    End Set
  End Property

  Private sType As String
  Private sInsUpdType As String
  Property Type__String() As String
    Get
      Return sType
    End Get
    Set(ByVal Value As String)
      sType = Value
      sInsUpdType = oUtil.FixParam(sType, True)
    End Set
  End Property

  Private dAmount As Double
  Private sInsUpdAmount As String
  Property Amount__Numeric() As Double
    Get
      Return dAmount
    End Get
    Set(ByVal Value As Double)
      dAmount = Value
      sInsUpdAmount = oUtil.FixParam(dAmount, True)
    End Set
  End Property

  Private sDate As String
  Private sInsUpdDate As String
  Property Date__Date() As String
    Get
      Return sDate
    End Get
    Set(ByVal Value As String)
      sDate = Value
      sInsUpdDate = oUtil.FixParam(sDate, True)
    End Set
  End Property

  Private sComputer As String
  Private sInsUpdComputer As String
  Property Computer__String() As String
    Get
      Return sComputer
    End Get
    Set(ByVal Value As String)
      sComputer = Value
      sInsUpdComputer = oUtil.FixParam(sComputer, True)
    End Set
  End Property

  Private sCheckNumber As String
  Private sInsUpdCheckNumber As String
  Property CheckNumber__String() As String
    Get
      Return sCheckNumber
    End Get
    Set(ByVal Value As String)
      sCheckNumber = Value
      sInsUpdCheckNumber = oUtil.FixParam(sCheckNumber, True)
    End Set
  End Property

  Private sCheckName As String
  Private sInsUpdCheckName As String
  Property CheckName__String() As String
    Get
      Return sCheckName
    End Get
    Set(ByVal Value As String)
      sCheckName = Value
      sInsUpdCheckName = oUtil.FixParam(sCheckName, True)
    End Set
  End Property

  Private sCCType As String
  Private sInsUpdCCType As String
  Property CCType__String() As String
    Get
      Return sCCType
    End Get
    Set(ByVal Value As String)
      sCCType = Value
      sInsUpdCCType = oUtil.FixParam(sCCType, True)
    End Set
  End Property

  Private sCCNumber As String
  Private sInsUpdCCNumber As String
  Property CCNumber__String() As String
    Get
      Return sCCNumber
    End Get
    Set(ByVal Value As String)
      sCCNumber = Value
      sInsUpdCCNumber = oUtil.FixParam(sCCNumber, True)
    End Set
  End Property

  Private sCCNumberEncrypted As String
  Private sInsUpdCCNumberEncrypted As String
  Property CCNumberEncrypted__String() As String
    Get
      Return sCCNumberEncrypted
    End Get
    Set(ByVal Value As String)
      sCCNumberEncrypted = Value
      sInsUpdCCNumberEncrypted = oUtil.FixParam(sCCNumberEncrypted, True)
    End Set
  End Property

  Private sCCName As String
  Private sInsUpdCCName As String
  Property CCName__String() As String
    Get
      Return sCCName
    End Get
    Set(ByVal Value As String)
      sCCName = Value
      sInsUpdCCName = oUtil.FixParam(sCCName, True)
    End Set
  End Property

  Private sCCExpMonth As String
  Private sInsUpdCCExpMonth As String
  Property CCExpMonth__String() As String
    Get
      Return sCCExpMonth
    End Get
    Set(ByVal Value As String)
      sCCExpMonth = Value
      sInsUpdCCExpMonth = oUtil.FixParam(sCCExpMonth, True)
    End Set
  End Property

  Private sCCExpYear As String
  Private sInsUpdCCExpYear As String
  Property CCExpYear__String() As String
    Get
      Return sCCExpYear
    End Get
    Set(ByVal Value As String)
      sCCExpYear = Value
      sInsUpdCCExpYear = oUtil.FixParam(sCCExpYear, True)
    End Set
  End Property

  Private sCCVerification As String
  Private sInsUpdCCVerification As String
  Property CCVerification__String() As String
    Get
      Return sCCVerification
    End Get
    Set(ByVal Value As String)
      sCCVerification = Value
      sInsUpdCCVerification = oUtil.FixParam(sCCVerification, True)
    End Set
  End Property

  Private sCCConfirm As String
  Private sInsUpdCCConfirm As String
  Property CCConfirm__String() As String
    Get
      Return sCCConfirm
    End Get
    Set(ByVal Value As String)
      sCCConfirm = Value
      sInsUpdCCConfirm = oUtil.FixParam(sCCConfirm, True)
    End Set
  End Property

  Private sQBStatus As String
  Private sInsUpdQBStatus As String
  Property QBStatus__String() As String
    Get
      Return sQBStatus
    End Get
    Set(ByVal Value As String)
      sQBStatus = Value
      sInsUpdQBStatus = oUtil.FixParam(sQBStatus, True)
    End Set
  End Property

  Private sNotes As String
  Private sInsUpdNotes As String
  Property Notes__String() As String
    Get
      Return sNotes
    End Get
    Set(ByVal Value As String)
      sNotes = Value
      sInsUpdNotes = oUtil.FixParam(sNotes, True)
    End Set
  End Property

  Private sUser As String
  Private sInsUpdUser As String
  Property User__String() As String
    Get
      Return sUser
    End Get
    Set(ByVal Value As String)
      sUser = Value
      sInsUpdUser = oUtil.FixParam(sUser, True)
    End Set
  End Property

  Private sStatus As String
  Private sInsUpdStatus As String
  Property Status__String() As String
    Get
      Return sStatus
    End Get
    Set(ByVal Value As String)
      sStatus = Value
      sInsUpdStatus = oUtil.FixParam(sStatus, True)
    End Set
  End Property

  Private iUseOnReceipt As Int32
  Private sInsUpdUseOnReceipt As String
  Property UseOnReceipt__Integer() As Int32
    Get
      Return iUseOnReceipt
    End Get
    Set(ByVal Value As Int32)
      iUseOnReceipt = Value
      sInsUpdUseOnReceipt = oUtil.FixParam(iUseOnReceipt, True)
    End Set
  End Property

  Private sOrigCCNumber As String
  Private sInsUpdOrigCCNumber As String
  Property OrigCCNumber__String() As String
    Get
      Return sOrigCCNumber
    End Get
    Set(ByVal Value As String)
      sOrigCCNumber = Value
      sInsUpdOrigCCNumber = oUtil.FixParam(sOrigCCNumber, True)
    End Set
  End Property

  Private sCCAddress As String
  Private sInsUpdCCAddress As String
  Property CCAddress__String() As String
    Get
      Return sCCAddress
    End Get
    Set(ByVal Value As String)
      sCCAddress = Value
      sInsUpdCCAddress = oUtil.FixParam(sCCAddress, True)
    End Set
  End Property

  Private sCCCity As String
  Private sInsUpdCCCity As String
  Property CCCity__String() As String
    Get
      Return sCCCity
    End Get
    Set(ByVal Value As String)
      sCCCity = Value
      sInsUpdCCCity = oUtil.FixParam(sCCCity, True)
    End Set
  End Property

  Private sCCState As String
  Private sInsUpdCCState As String
  Property CCState__String() As String
    Get
      Return sCCState
    End Get
    Set(ByVal Value As String)
      sCCState = Value
      sInsUpdCCState = oUtil.FixParam(sCCState, True)
    End Set
  End Property

  Private sCCZip As String
  Private sInsUpdCCZip As String
  Property CCZip__String() As String
    Get
      Return sCCZip
    End Get
    Set(ByVal Value As String)
      sCCZip = Value
      sInsUpdCCZip = oUtil.FixParam(sCCZip, True)
    End Set
  End Property

  Private sAuthorizeNetReturnCode As String
  Private sInsUpdAuthorizeNetReturnCode As String
  Property AuthorizeNetReturnCode__String() As String
    Get
      Return sAuthorizeNetReturnCode
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetReturnCode = Value
      sInsUpdAuthorizeNetReturnCode = oUtil.FixParam(sAuthorizeNetReturnCode, True)
    End Set
  End Property

  Private sAuthorizeNetReasonCode As String
  Private sInsUpdAuthorizeNetReasonCode As String
  Property AuthorizeNetReasonCode__String() As String
    Get
      Return sAuthorizeNetReasonCode
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetReasonCode = Value
      sInsUpdAuthorizeNetReasonCode = oUtil.FixParam(sAuthorizeNetReasonCode, True)
    End Set
  End Property

  Private sAuthorizeNetReasonText As String
  Private sInsUpdAuthorizeNetReasonText As String
  Property AuthorizeNetReasonText__String() As String
    Get
      Return sAuthorizeNetReasonText
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetReasonText = Value
      sInsUpdAuthorizeNetReasonText = oUtil.FixParam(sAuthorizeNetReasonText, True)
    End Set
  End Property

  Private sAuthorizeNetApprovalCode As String
  Private sInsUpdAuthorizeNetApprovalCode As String
  Property AuthorizeNetApprovalCode__String() As String
    Get
      Return sAuthorizeNetApprovalCode
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetApprovalCode = Value
      sInsUpdAuthorizeNetApprovalCode = oUtil.FixParam(sAuthorizeNetApprovalCode, True)
    End Set
  End Property

  Private sAuthorizeNetAVSResultCode As String
  Private sInsUpdAuthorizeNetAVSResultCode As String
  Property AuthorizeNetAVSResultCode__String() As String
    Get
      Return sAuthorizeNetAVSResultCode
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetAVSResultCode = Value
      sInsUpdAuthorizeNetAVSResultCode = oUtil.FixParam(sAuthorizeNetAVSResultCode, True)
    End Set
  End Property

  Private sAuthorizeNetAVSResultText As String
  Private sInsUpdAuthorizeNetAVSResultText As String
  Property AuthorizeNetAVSResultText__String() As String
    Get
      Return sAuthorizeNetAVSResultText
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetAVSResultText = Value
      sInsUpdAuthorizeNetAVSResultText = oUtil.FixParam(sAuthorizeNetAVSResultText, True)
    End Set
  End Property

  Private sAuthorizeNetTransactionID As String
  Private sInsUpdAuthorizeNetTransactionID As String
  Property AuthorizeNetTransactionID__String() As String
    Get
      Return sAuthorizeNetTransactionID
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetTransactionID = Value
      sInsUpdAuthorizeNetTransactionID = oUtil.FixParam(sAuthorizeNetTransactionID, True)
    End Set
  End Property

  Private sAuthorizeNetCVCCResponseCode As String
  Private sInsUpdAuthorizeNetCVCCResponseCode As String
  Property AuthorizeNetCVCCResponseCode__String() As String
    Get
      Return sAuthorizeNetCVCCResponseCode
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetCVCCResponseCode = Value
      sInsUpdAuthorizeNetCVCCResponseCode = oUtil.FixParam(sAuthorizeNetCVCCResponseCode, True)
    End Set
  End Property

  Private sAuthorizeNetCVCCResponseText As String
  Private sInsUpdAuthorizeNetCVCCResponseText As String
  Property AuthorizeNetCVCCResponseText__String() As String
    Get
      Return sAuthorizeNetCVCCResponseText
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetCVCCResponseText = Value
      sInsUpdAuthorizeNetCVCCResponseText = oUtil.FixParam(sAuthorizeNetCVCCResponseText, True)
    End Set
  End Property

  Private sLocation As String
  Private sInsUpdLocation As String
  Property Location__String() As String
    Get
      Return sLocation
    End Get
    Set(ByVal Value As String)
      sLocation = Value
      sInsUpdLocation = oUtil.FixParam(sLocation, True)
    End Set
  End Property

  Private sEmailSent As String
  Private sInsUpdEmailSent As String
  Property EmailSent__String() As String
    Get
      Return sEmailSent
    End Get
    Set(ByVal Value As String)
      sEmailSent = Value
      sInsUpdEmailSent = oUtil.FixParam(sEmailSent, True)
    End Set
  End Property

  Private sOriginalPaymentCategory As String
  Private sInsUpdOriginalPaymentCategory As String
  Property OriginalPaymentCategory__String() As String
    Get
      Return sOriginalPaymentCategory
    End Get
    Set(ByVal Value As String)
      sOriginalPaymentCategory = Value
      sInsUpdOriginalPaymentCategory = oUtil.FixParam(sOriginalPaymentCategory, True)
    End Set
  End Property

  Private iOriginalPaymentID As Int32
  Private sInsUpdOriginalPaymentID As String
  Property OriginalPaymentID__Integer() As Int32
    Get
      Return iOriginalPaymentID
    End Get
    Set(ByVal Value As Int32)
      iOriginalPaymentID = Value
      sInsUpdOriginalPaymentID = oUtil.FixParam(iOriginalPaymentID, True)
    End Set
  End Property

  Private sBankName As String
  Private sInsUpdBankName As String
  Property BankName__String() As String
    Get
      Return sBankName
    End Get
    Set(ByVal Value As String)
      sBankName = Value
      sInsUpdBankName = oUtil.FixParam(sBankName, True)
    End Set
  End Property

  Private sBankAccountType As String
  Private sInsUpdBankAccountType As String
  Property BankAccountType__String() As String
    Get
      Return sBankAccountType
    End Get
    Set(ByVal Value As String)
      sBankAccountType = Value
      sInsUpdBankAccountType = oUtil.FixParam(sBankAccountType, True)
    End Set
  End Property

  Private sBankAccountName As String
  Private sInsUpdBankAccountName As String
  Property BankAccountName__String() As String
    Get
      Return sBankAccountName
    End Get
    Set(ByVal Value As String)
      sBankAccountName = Value
      sInsUpdBankAccountName = oUtil.FixParam(sBankAccountName, True)
    End Set
  End Property

  Private sBankAccountNumber As String
  Private sInsUpdBankAccountNumber As String
  Property BankAccountNumber__String() As String
    Get
      Return sBankAccountNumber
    End Get
    Set(ByVal Value As String)
      sBankAccountNumber = Value
      sInsUpdBankAccountNumber = oUtil.FixParam(sBankAccountNumber, True)
    End Set
  End Property

  Private sBankAccountNumberEncrypted As String
  Private sInsUpdBankAccountNumberEncrypted As String
  Property BankAccountNumberEncrypted__String() As String
    Get
      Return sBankAccountNumberEncrypted
    End Get
    Set(ByVal Value As String)
      sBankAccountNumberEncrypted = Value
      sInsUpdBankAccountNumberEncrypted = oUtil.FixParam(sBankAccountNumberEncrypted, True)
    End Set
  End Property

  Private sBankRoutingNumber As String
  Private sInsUpdBankRoutingNumber As String
  Property BankRoutingNumber__String() As String
    Get
      Return sBankRoutingNumber
    End Get
    Set(ByVal Value As String)
      sBankRoutingNumber = Value
      sInsUpdBankRoutingNumber = oUtil.FixParam(sBankRoutingNumber, True)
    End Set
  End Property

  Private sEmail As String
  Private sInsUpdEmail As String
  Property Email__String() As String
    Get
      Return sEmail
    End Get
    Set(ByVal Value As String)
      sEmail = Value
      sInsUpdEmail = oUtil.FixParam(sEmail, True)
    End Set
  End Property

  Private sAuthorizeNetURL As String
  Private sInsUpdAuthorizeNetURL As String
  Property AuthorizeNetURL__String() As String
    Get
      Return sAuthorizeNetURL
    End Get
    Set(ByVal Value As String)
      sAuthorizeNetURL = Value
      sInsUpdAuthorizeNetURL = oUtil.FixParam(sAuthorizeNetURL, True)
    End Set
  End Property

  Private iRefundPaymentID As Int32
  Private sInsUpdRefundPaymentID As String
  Property RefundPaymentID__Integer() As Int32
    Get
      Return iRefundPaymentID
    End Get
    Set(ByVal Value As Int32)
      iRefundPaymentID = Value
      sInsUpdRefundPaymentID = oUtil.FixParam(iRefundPaymentID, True)
    End Set
  End Property

  Private sRefundPaymentTable As String
  Private sInsUpdRefundPaymentTable As String
  Property RefundPaymentTable__String() As String
    Get
      Return sRefundPaymentTable
    End Get
    Set(ByVal Value As String)
      sRefundPaymentTable = Value
      sInsUpdRefundPaymentTable = oUtil.FixParam(sRefundPaymentTable, True)
    End Set
  End Property

  Private sCheckAddress As String
  Private sInsUpdCheckAddress As String
  Property CheckAddress__String() As String
    Get
      Return sCheckAddress
    End Get
    Set(ByVal Value As String)
      sCheckAddress = Value
      sInsUpdCheckAddress = oUtil.FixParam(sCheckAddress, True)
    End Set
  End Property

  Private sCheckCity As String
  Private sInsUpdCheckCity As String
  Property CheckCity__String() As String
    Get
      Return sCheckCity
    End Get
    Set(ByVal Value As String)
      sCheckCity = Value
      sInsUpdCheckCity = oUtil.FixParam(sCheckCity, True)
    End Set
  End Property

  Private sCheckState As String
  Private sInsUpdCheckState As String
  Property CheckState__String() As String
    Get
      Return sCheckState
    End Get
    Set(ByVal Value As String)
      sCheckState = Value
      sInsUpdCheckState = oUtil.FixParam(sCheckState, True)
    End Set
  End Property

  Private sCheckZip As String
  Private sInsUpdCheckZip As String
  Property CheckZip__String() As String
    Get
      Return sCheckZip
    End Get
    Set(ByVal Value As String)
      sCheckZip = Value
      sInsUpdCheckZip = oUtil.FixParam(sCheckZip, True)
    End Set
  End Property

  Private sQBListID As String
  Private sInsUpdQBListID As String
  Property QBListID__String() As String
    Get
      Return sQBListID
    End Get
    Set(ByVal Value As String)
      sQBListID = Value
      sInsUpdQBListID = oUtil.FixParam(sQBListID, True)
    End Set
  End Property

  Private iQBActivityID As Int32
  Private sInsUpdQBActivityID As String
  Property QBActivityID__Integer() As Int32
    Get
      Return iQBActivityID
    End Get
    Set(ByVal Value As Int32)
      iQBActivityID = Value
      sInsUpdQBActivityID = oUtil.FixParam(iQBActivityID, True)
    End Set
  End Property

  Private sAuthNETPaymentProfileID As String
  Private sInsUpdAuthNETPaymentProfileID As String
  Property AuthNETPaymentProfileID__String() As String
    Get
      Return sAuthNETPaymentProfileID
    End Get
    Set(ByVal Value As String)
      sAuthNETPaymentProfileID = Value
      sInsUpdAuthNETPaymentProfileID = oUtil.FixParam(sAuthNETPaymentProfileID, True)
    End Set
  End Property

  Public Sub Clear()
    iPaymentID = 0
    sInsUpdPaymentID = ""
    iBookingID = 0
    sInsUpdBookingID = ""
    iPropertyID = 0
    sInsUpdPropertyID = ""
    iHostID = 0
    sInsUpdHostID = ""
    iGuestID = 0
    sInsUpdGuestID = ""
    sCategory = ""
    sInsUpdCategory = ""
    sType = ""
    sInsUpdType = ""
    dAmount = 0.0
    sInsUpdAmount = ""
    sDate = ""
    sInsUpdDate = ""
    sComputer = ""
    sInsUpdComputer = ""
    sCheckNumber = ""
    sInsUpdCheckNumber = ""
    sCheckName = ""
    sInsUpdCheckName = ""
    sCCType = ""
    sInsUpdCCType = ""
    sCCNumber = ""
    sInsUpdCCNumber = ""
    sCCNumberEncrypted = ""
    sInsUpdCCNumberEncrypted = ""
    sCCName = ""
    sInsUpdCCName = ""
    sCCExpMonth = ""
    sInsUpdCCExpMonth = ""
    sCCExpYear = ""
    sInsUpdCCExpYear = ""
    sCCVerification = ""
    sInsUpdCCVerification = ""
    sCCConfirm = ""
    sInsUpdCCConfirm = ""
    sQBStatus = ""
    sInsUpdQBStatus = ""
    sNotes = ""
    sInsUpdNotes = ""
    sUser = ""
    sInsUpdUser = ""
    sStatus = ""
    sInsUpdStatus = ""
    iUseOnReceipt = 0
    sInsUpdUseOnReceipt = ""
    sOrigCCNumber = ""
    sInsUpdOrigCCNumber = ""
    sCCAddress = ""
    sInsUpdCCAddress = ""
    sCCCity = ""
    sInsUpdCCCity = ""
    sCCState = ""
    sInsUpdCCState = ""
    sCCZip = ""
    sInsUpdCCZip = ""
    sAuthorizeNetReturnCode = ""
    sInsUpdAuthorizeNetReturnCode = ""
    sAuthorizeNetReasonCode = ""
    sInsUpdAuthorizeNetReasonCode = ""
    sAuthorizeNetReasonText = ""
    sInsUpdAuthorizeNetReasonText = ""
    sAuthorizeNetApprovalCode = ""
    sInsUpdAuthorizeNetApprovalCode = ""
    sAuthorizeNetAVSResultCode = ""
    sInsUpdAuthorizeNetAVSResultCode = ""
    sAuthorizeNetAVSResultText = ""
    sInsUpdAuthorizeNetAVSResultText = ""
    sAuthorizeNetTransactionID = ""
    sInsUpdAuthorizeNetTransactionID = ""
    sAuthorizeNetCVCCResponseCode = ""
    sInsUpdAuthorizeNetCVCCResponseCode = ""
    sAuthorizeNetCVCCResponseText = ""
    sInsUpdAuthorizeNetCVCCResponseText = ""
    sLocation = ""
    sInsUpdLocation = ""
    sEmailSent = ""
    sInsUpdEmailSent = ""
    sOriginalPaymentCategory = ""
    sInsUpdOriginalPaymentCategory = ""
    iOriginalPaymentID = 0
    sInsUpdOriginalPaymentID = ""
    sBankName = ""
    sInsUpdBankName = ""
    sBankAccountType = ""
    sInsUpdBankAccountType = ""
    sBankAccountName = ""
    sInsUpdBankAccountName = ""
    sBankAccountNumber = ""
    sInsUpdBankAccountNumber = ""
    sBankAccountNumberEncrypted = ""
    sInsUpdBankAccountNumberEncrypted = ""
    sBankRoutingNumber = ""
    sInsUpdBankRoutingNumber = ""
    sEmail = ""
    sInsUpdEmail = ""
    sAuthorizeNetURL = ""
    sInsUpdAuthorizeNetURL = ""
    iRefundPaymentID = 0
    sInsUpdRefundPaymentID = ""
    sRefundPaymentTable = ""
    sInsUpdRefundPaymentTable = ""
    sCheckAddress = ""
    sInsUpdCheckAddress = ""
    sCheckCity = ""
    sInsUpdCheckCity = ""
    sCheckState = ""
    sInsUpdCheckState = ""
    sCheckZip = ""
    sInsUpdCheckZip = ""
    sQBListID = ""
    sInsUpdQBListID = ""
    iQBActivityID = 0
    sInsUpdQBActivityID = ""
    sAuthNETPaymentProfileID = ""
    sInsUpdAuthNETPaymentProfileID = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdPaymentID.ToString = "'-12345'" Then sb.Append("PaymentID,")
        If sInsUpdBookingID.ToString = "'-12345'" Then sb.Append("BookingID,")
        If sInsUpdPropertyID.ToString = "'-12345'" Then sb.Append("PropertyID,")
        If sInsUpdHostID.ToString = "'-12345'" Then sb.Append("HostID,")
        If sInsUpdGuestID.ToString = "'-12345'" Then sb.Append("GuestID,")
        If sInsUpdCategory.ToString = "'Select'" Then sb.Append("Category,")
        If sInsUpdType.ToString = "'Select'" Then sb.Append("Type,")
        If sInsUpdAmount.ToString = "'-12345'" Then sb.Append("Amount,")
        If sInsUpdDate.ToString = "'Select'" Then sb.Append("[Date],")
        If sInsUpdComputer.ToString = "'Select'" Then sb.Append("Computer,")
        If sInsUpdCheckNumber.ToString = "'Select'" Then sb.Append("CheckNumber,")
        If sInsUpdCheckName.ToString = "'Select'" Then sb.Append("CheckName,")
        If sInsUpdCCType.ToString = "'Select'" Then sb.Append("CCType,")
        If sInsUpdCCNumber.ToString = "'Select'" Then sb.Append("CCNumber,")
        If sInsUpdCCNumberEncrypted.ToString = "'Select'" Then sb.Append("CCNumberEncrypted,")
        If sInsUpdCCName.ToString = "'Select'" Then sb.Append("CCName,")
        If sInsUpdCCExpMonth.ToString = "'Select'" Then sb.Append("CCExpMonth,")
        If sInsUpdCCExpYear.ToString = "'Select'" Then sb.Append("CCExpYear,")
        If sInsUpdCCVerification.ToString = "'Select'" Then sb.Append("CCVerification,")
        If sInsUpdCCConfirm.ToString = "'Select'" Then sb.Append("CCConfirm,")
        If sInsUpdQBStatus.ToString = "'Select'" Then sb.Append("QBStatus,")
        If sInsUpdNotes.ToString = "'Select'" Then sb.Append("Notes,")
        If sInsUpdUser.ToString = "'Select'" Then sb.Append("[User],")
        If sInsUpdStatus.ToString = "'Select'" Then sb.Append("Status,")
        If sInsUpdUseOnReceipt.ToString = "'-12345'" Then sb.Append("UseOnReceipt,")
        If sInsUpdOrigCCNumber.ToString = "'Select'" Then sb.Append("OrigCCNumber,")
        If sInsUpdCCAddress.ToString = "'Select'" Then sb.Append("CCAddress,")
        If sInsUpdCCCity.ToString = "'Select'" Then sb.Append("CCCity,")
        If sInsUpdCCState.ToString = "'Select'" Then sb.Append("CCState,")
        If sInsUpdCCZip.ToString = "'Select'" Then sb.Append("CCZip,")
        If sInsUpdAuthorizeNetReturnCode.ToString = "'Select'" Then sb.Append("AuthorizeNetReturnCode,")
        If sInsUpdAuthorizeNetReasonCode.ToString = "'Select'" Then sb.Append("AuthorizeNetReasonCode,")
        If sInsUpdAuthorizeNetReasonText.ToString = "'Select'" Then sb.Append("AuthorizeNetReasonText,")
        If sInsUpdAuthorizeNetApprovalCode.ToString = "'Select'" Then sb.Append("AuthorizeNetApprovalCode,")
        If sInsUpdAuthorizeNetAVSResultCode.ToString = "'Select'" Then sb.Append("AuthorizeNetAVSResultCode,")
        If sInsUpdAuthorizeNetAVSResultText.ToString = "'Select'" Then sb.Append("AuthorizeNetAVSResultText,")
        If sInsUpdAuthorizeNetTransactionID.ToString = "'Select'" Then sb.Append("AuthorizeNetTransactionID,")
        If sInsUpdAuthorizeNetCVCCResponseCode.ToString = "'Select'" Then sb.Append("AuthorizeNetCVCCResponseCode,")
        If sInsUpdAuthorizeNetCVCCResponseText.ToString = "'Select'" Then sb.Append("AuthorizeNetCVCCResponseText,")
        If sInsUpdLocation.ToString = "'Select'" Then sb.Append("Location,")
        If sInsUpdEmailSent.ToString = "'Select'" Then sb.Append("EmailSent,")
        If sInsUpdOriginalPaymentCategory.ToString = "'Select'" Then sb.Append("OriginalPaymentCategory,")
        If sInsUpdOriginalPaymentID.ToString = "'-12345'" Then sb.Append("OriginalPaymentID,")
        If sInsUpdBankName.ToString = "'Select'" Then sb.Append("BankName,")
        If sInsUpdBankAccountType.ToString = "'Select'" Then sb.Append("BankAccountType,")
        If sInsUpdBankAccountName.ToString = "'Select'" Then sb.Append("BankAccountName,")
        If sInsUpdBankAccountNumber.ToString = "'Select'" Then sb.Append("BankAccountNumber,")
        If sInsUpdBankAccountNumberEncrypted.ToString = "'Select'" Then sb.Append("BankAccountNumberEncrypted,")
        If sInsUpdBankRoutingNumber.ToString = "'Select'" Then sb.Append("BankRoutingNumber,")
        If sInsUpdEmail.ToString = "'Select'" Then sb.Append("Email,")
        If sInsUpdAuthorizeNetURL.ToString = "'Select'" Then sb.Append("AuthorizeNetURL,")
        If sInsUpdRefundPaymentID.ToString = "'-12345'" Then sb.Append("RefundPaymentID,")
        If sInsUpdRefundPaymentTable.ToString = "'Select'" Then sb.Append("RefundPaymentTable,")
        If sInsUpdCheckAddress.ToString = "'Select'" Then sb.Append("CheckAddress,")
        If sInsUpdCheckCity.ToString = "'Select'" Then sb.Append("CheckCity,")
        If sInsUpdCheckState.ToString = "'Select'" Then sb.Append("CheckState,")
        If sInsUpdCheckZip.ToString = "'Select'" Then sb.Append("CheckZip,")
        If sInsUpdQBListID.ToString = "'Select'" Then sb.Append("QBListID,")
        If sInsUpdQBActivityID.ToString = "'-12345'" Then sb.Append("QBActivityID,")
        If sInsUpdAuthNETPaymentProfileID.ToString = "'Select'" Then sb.Append("AuthNETPaymentProfileID,")
      Else
        sb.Append("PaymentID,")
        sb.Append("BookingID,")
        sb.Append("PropertyID,")
        sb.Append("HostID,")
        sb.Append("GuestID,")
        sb.Append("Category,")
        sb.Append("Type,")
        sb.Append("Amount,")
        sb.Append("[Date],")
        sb.Append("Computer,")
        sb.Append("CheckNumber,")
        sb.Append("CheckName,")
        sb.Append("CCType,")
        sb.Append("CCNumber,")
        sb.Append("CCNumberEncrypted,")
        sb.Append("CCName,")
        sb.Append("CCExpMonth,")
        sb.Append("CCExpYear,")
        sb.Append("CCVerification,")
        sb.Append("CCConfirm,")
        sb.Append("QBStatus,")
        sb.Append("Notes,")
        sb.Append("[User],")
        sb.Append("Status,")
        sb.Append("UseOnReceipt,")
        sb.Append("OrigCCNumber,")
        sb.Append("CCAddress,")
        sb.Append("CCCity,")
        sb.Append("CCState,")
        sb.Append("CCZip,")
        sb.Append("AuthorizeNetReturnCode,")
        sb.Append("AuthorizeNetReasonCode,")
        sb.Append("AuthorizeNetReasonText,")
        sb.Append("AuthorizeNetApprovalCode,")
        sb.Append("AuthorizeNetAVSResultCode,")
        sb.Append("AuthorizeNetAVSResultText,")
        sb.Append("AuthorizeNetTransactionID,")
        sb.Append("AuthorizeNetCVCCResponseCode,")
        sb.Append("AuthorizeNetCVCCResponseText,")
        sb.Append("Location,")
        sb.Append("EmailSent,")
        sb.Append("OriginalPaymentCategory,")
        sb.Append("OriginalPaymentID,")
        sb.Append("BankName,")
        sb.Append("BankAccountType,")
        sb.Append("BankAccountName,")
        sb.Append("BankAccountNumber,")
        sb.Append("BankAccountNumberEncrypted,")
        sb.Append("BankRoutingNumber,")
        sb.Append("Email,")
        sb.Append("AuthorizeNetURL,")
        sb.Append("RefundPaymentID,")
        sb.Append("RefundPaymentTable,")
        sb.Append("CheckAddress,")
        sb.Append("CheckCity,")
        sb.Append("CheckState,")
        sb.Append("CheckZip,")
        sb.Append("QBListID,")
        sb.Append("QBActivityID,")
        sb.Append("AuthNETPaymentProfileID,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [RefundPayments]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdPaymentID.ToString <> "") And (sInsUpdPaymentID <> "'-12345'") Then sbw.Append("PaymentID=" & sInsUpdPaymentID & " and ")
      If (sInsUpdBookingID.ToString <> "") And (sInsUpdBookingID <> "'-12345'") Then sbw.Append("BookingID=" & sInsUpdBookingID & " and ")
      If (sInsUpdPropertyID.ToString <> "") And (sInsUpdPropertyID <> "'-12345'") Then sbw.Append("PropertyID=" & sInsUpdPropertyID & " and ")
      If (sInsUpdHostID.ToString <> "") And (sInsUpdHostID <> "'-12345'") Then sbw.Append("HostID=" & sInsUpdHostID & " and ")
      If (sInsUpdGuestID.ToString <> "") And (sInsUpdGuestID <> "'-12345'") Then sbw.Append("GuestID=" & sInsUpdGuestID & " and ")
      If (sInsUpdCategory.ToString <> "") And (sInsUpdCategory <> "'Select'") Then sbw.Append("Category=" & sInsUpdCategory & " and ")
      If (sInsUpdType.ToString <> "") And (sInsUpdType <> "'Select'") Then sbw.Append("Type=" & sInsUpdType & " and ")
      If (sInsUpdAmount.ToString <> "") And (sInsUpdAmount <> "'-12345'") Then sbw.Append("Amount=" & sInsUpdAmount & " and ")
      If (sInsUpdDate.ToString <> "") And (sInsUpdDate <> "'Select'") Then sbw.Append("[Date]=" & sInsUpdDate & " and ")
      If (sInsUpdComputer.ToString <> "") And (sInsUpdComputer <> "'Select'") Then sbw.Append("Computer=" & sInsUpdComputer & " and ")
      If (sInsUpdCheckNumber.ToString <> "") And (sInsUpdCheckNumber <> "'Select'") Then sbw.Append("CheckNumber=" & sInsUpdCheckNumber & " and ")
      If (sInsUpdCheckName.ToString <> "") And (sInsUpdCheckName <> "'Select'") Then sbw.Append("CheckName=" & sInsUpdCheckName & " and ")
      If (sInsUpdCCType.ToString <> "") And (sInsUpdCCType <> "'Select'") Then sbw.Append("CCType=" & sInsUpdCCType & " and ")
      If (sInsUpdCCNumber.ToString <> "") And (sInsUpdCCNumber <> "'Select'") Then sbw.Append("CCNumber=" & sInsUpdCCNumber & " and ")
      If (sInsUpdCCNumberEncrypted.ToString <> "") And (sInsUpdCCNumberEncrypted <> "'Select'") Then sbw.Append("CCNumberEncrypted=" & sInsUpdCCNumberEncrypted & " and ")
      If (sInsUpdCCName.ToString <> "") And (sInsUpdCCName <> "'Select'") Then sbw.Append("CCName=" & sInsUpdCCName & " and ")
      If (sInsUpdCCExpMonth.ToString <> "") And (sInsUpdCCExpMonth <> "'Select'") Then sbw.Append("CCExpMonth=" & sInsUpdCCExpMonth & " and ")
      If (sInsUpdCCExpYear.ToString <> "") And (sInsUpdCCExpYear <> "'Select'") Then sbw.Append("CCExpYear=" & sInsUpdCCExpYear & " and ")
      If (sInsUpdCCVerification.ToString <> "") And (sInsUpdCCVerification <> "'Select'") Then sbw.Append("CCVerification=" & sInsUpdCCVerification & " and ")
      If (sInsUpdCCConfirm.ToString <> "") And (sInsUpdCCConfirm <> "'Select'") Then sbw.Append("CCConfirm=" & sInsUpdCCConfirm & " and ")
      If (sInsUpdQBStatus.ToString <> "") And (sInsUpdQBStatus <> "'Select'") Then sbw.Append("QBStatus=" & sInsUpdQBStatus & " and ")
      If (sInsUpdNotes.ToString <> "") And (sInsUpdNotes <> "'Select'") Then sbw.Append("Notes=" & sInsUpdNotes & " and ")
      If (sInsUpdUser.ToString <> "") And (sInsUpdUser <> "'Select'") Then sbw.Append("[User]=" & sInsUpdUser & " and ")
      If (sInsUpdStatus.ToString <> "") And (sInsUpdStatus <> "'Select'") Then sbw.Append("Status=" & sInsUpdStatus & " and ")
      If (sInsUpdUseOnReceipt.ToString <> "") And (sInsUpdUseOnReceipt <> "'-12345'") Then sbw.Append("UseOnReceipt=" & sInsUpdUseOnReceipt & " and ")
      If (sInsUpdOrigCCNumber.ToString <> "") And (sInsUpdOrigCCNumber <> "'Select'") Then sbw.Append("OrigCCNumber=" & sInsUpdOrigCCNumber & " and ")
      If (sInsUpdCCAddress.ToString <> "") And (sInsUpdCCAddress <> "'Select'") Then sbw.Append("CCAddress=" & sInsUpdCCAddress & " and ")
      If (sInsUpdCCCity.ToString <> "") And (sInsUpdCCCity <> "'Select'") Then sbw.Append("CCCity=" & sInsUpdCCCity & " and ")
      If (sInsUpdCCState.ToString <> "") And (sInsUpdCCState <> "'Select'") Then sbw.Append("CCState=" & sInsUpdCCState & " and ")
      If (sInsUpdCCZip.ToString <> "") And (sInsUpdCCZip <> "'Select'") Then sbw.Append("CCZip=" & sInsUpdCCZip & " and ")
      If (sInsUpdAuthorizeNetReturnCode.ToString <> "") And (sInsUpdAuthorizeNetReturnCode <> "'Select'") Then sbw.Append("AuthorizeNetReturnCode=" & sInsUpdAuthorizeNetReturnCode & " and ")
      If (sInsUpdAuthorizeNetReasonCode.ToString <> "") And (sInsUpdAuthorizeNetReasonCode <> "'Select'") Then sbw.Append("AuthorizeNetReasonCode=" & sInsUpdAuthorizeNetReasonCode & " and ")
      If (sInsUpdAuthorizeNetReasonText.ToString <> "") And (sInsUpdAuthorizeNetReasonText <> "'Select'") Then sbw.Append("AuthorizeNetReasonText=" & sInsUpdAuthorizeNetReasonText & " and ")
      If (sInsUpdAuthorizeNetApprovalCode.ToString <> "") And (sInsUpdAuthorizeNetApprovalCode <> "'Select'") Then sbw.Append("AuthorizeNetApprovalCode=" & sInsUpdAuthorizeNetApprovalCode & " and ")
      If (sInsUpdAuthorizeNetAVSResultCode.ToString <> "") And (sInsUpdAuthorizeNetAVSResultCode <> "'Select'") Then sbw.Append("AuthorizeNetAVSResultCode=" & sInsUpdAuthorizeNetAVSResultCode & " and ")
      If (sInsUpdAuthorizeNetAVSResultText.ToString <> "") And (sInsUpdAuthorizeNetAVSResultText <> "'Select'") Then sbw.Append("AuthorizeNetAVSResultText=" & sInsUpdAuthorizeNetAVSResultText & " and ")
      If (sInsUpdAuthorizeNetTransactionID.ToString <> "") And (sInsUpdAuthorizeNetTransactionID <> "'Select'") Then sbw.Append("AuthorizeNetTransactionID=" & sInsUpdAuthorizeNetTransactionID & " and ")
      If (sInsUpdAuthorizeNetCVCCResponseCode.ToString <> "") And (sInsUpdAuthorizeNetCVCCResponseCode <> "'Select'") Then sbw.Append("AuthorizeNetCVCCResponseCode=" & sInsUpdAuthorizeNetCVCCResponseCode & " and ")
      If (sInsUpdAuthorizeNetCVCCResponseText.ToString <> "") And (sInsUpdAuthorizeNetCVCCResponseText <> "'Select'") Then sbw.Append("AuthorizeNetCVCCResponseText=" & sInsUpdAuthorizeNetCVCCResponseText & " and ")
      If (sInsUpdLocation.ToString <> "") And (sInsUpdLocation <> "'Select'") Then sbw.Append("Location=" & sInsUpdLocation & " and ")
      If (sInsUpdEmailSent.ToString <> "") And (sInsUpdEmailSent <> "'Select'") Then sbw.Append("EmailSent=" & sInsUpdEmailSent & " and ")
      If (sInsUpdOriginalPaymentCategory.ToString <> "") And (sInsUpdOriginalPaymentCategory <> "'Select'") Then sbw.Append("OriginalPaymentCategory=" & sInsUpdOriginalPaymentCategory & " and ")
      If (sInsUpdOriginalPaymentID.ToString <> "") And (sInsUpdOriginalPaymentID <> "'-12345'") Then sbw.Append("OriginalPaymentID=" & sInsUpdOriginalPaymentID & " and ")
      If (sInsUpdBankName.ToString <> "") And (sInsUpdBankName <> "'Select'") Then sbw.Append("BankName=" & sInsUpdBankName & " and ")
      If (sInsUpdBankAccountType.ToString <> "") And (sInsUpdBankAccountType <> "'Select'") Then sbw.Append("BankAccountType=" & sInsUpdBankAccountType & " and ")
      If (sInsUpdBankAccountName.ToString <> "") And (sInsUpdBankAccountName <> "'Select'") Then sbw.Append("BankAccountName=" & sInsUpdBankAccountName & " and ")
      If (sInsUpdBankAccountNumber.ToString <> "") And (sInsUpdBankAccountNumber <> "'Select'") Then sbw.Append("BankAccountNumber=" & sInsUpdBankAccountNumber & " and ")
      If (sInsUpdBankAccountNumberEncrypted.ToString <> "") And (sInsUpdBankAccountNumberEncrypted <> "'Select'") Then sbw.Append("BankAccountNumberEncrypted=" & sInsUpdBankAccountNumberEncrypted & " and ")
      If (sInsUpdBankRoutingNumber.ToString <> "") And (sInsUpdBankRoutingNumber <> "'Select'") Then sbw.Append("BankRoutingNumber=" & sInsUpdBankRoutingNumber & " and ")
      If (sInsUpdEmail.ToString <> "") And (sInsUpdEmail <> "'Select'") Then sbw.Append("Email=" & sInsUpdEmail & " and ")
      If (sInsUpdAuthorizeNetURL.ToString <> "") And (sInsUpdAuthorizeNetURL <> "'Select'") Then sbw.Append("AuthorizeNetURL=" & sInsUpdAuthorizeNetURL & " and ")
      If (sInsUpdRefundPaymentID.ToString <> "") And (sInsUpdRefundPaymentID <> "'-12345'") Then sbw.Append("RefundPaymentID=" & sInsUpdRefundPaymentID & " and ")
      If (sInsUpdRefundPaymentTable.ToString <> "") And (sInsUpdRefundPaymentTable <> "'Select'") Then sbw.Append("RefundPaymentTable=" & sInsUpdRefundPaymentTable & " and ")
      If (sInsUpdCheckAddress.ToString <> "") And (sInsUpdCheckAddress <> "'Select'") Then sbw.Append("CheckAddress=" & sInsUpdCheckAddress & " and ")
      If (sInsUpdCheckCity.ToString <> "") And (sInsUpdCheckCity <> "'Select'") Then sbw.Append("CheckCity=" & sInsUpdCheckCity & " and ")
      If (sInsUpdCheckState.ToString <> "") And (sInsUpdCheckState <> "'Select'") Then sbw.Append("CheckState=" & sInsUpdCheckState & " and ")
      If (sInsUpdCheckZip.ToString <> "") And (sInsUpdCheckZip <> "'Select'") Then sbw.Append("CheckZip=" & sInsUpdCheckZip & " and ")
      If (sInsUpdQBListID.ToString <> "") And (sInsUpdQBListID <> "'Select'") Then sbw.Append("QBListID=" & sInsUpdQBListID & " and ")
      If (sInsUpdQBActivityID.ToString <> "") And (sInsUpdQBActivityID <> "'-12345'") Then sbw.Append("QBActivityID=" & sInsUpdQBActivityID & " and ")
      If (sInsUpdAuthNETPaymentProfileID.ToString <> "") And (sInsUpdAuthNETPaymentProfileID <> "'Select'") Then sbw.Append("AuthNETPaymentProfileID=" & sInsUpdAuthNETPaymentProfileID & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      PaymentID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("PaymentID")), 0, CurrentRow.Item("PaymentID").ToString)
      BookingID__Integer = IIf(IsDBNull(CurrentRow.Item("BookingID")), 0, CurrentRow.Item("BookingID"))
      PropertyID__Integer = IIf(IsDBNull(CurrentRow.Item("PropertyID")), 0, CurrentRow.Item("PropertyID"))
      HostID__Integer = IIf(IsDBNull(CurrentRow.Item("HostID")), 0, CurrentRow.Item("HostID"))
      GuestID__Integer = IIf(IsDBNull(CurrentRow.Item("GuestID")), 0, CurrentRow.Item("GuestID"))
      Category__String = IIf(IsDBNull(CurrentRow.Item("Category")), "", CurrentRow.Item("Category"))
      Type__String = IIf(IsDBNull(CurrentRow.Item("Type")), "", CurrentRow.Item("Type"))
      Amount__Numeric = IIf(IsDBNull(CurrentRow.Item("Amount")), 0.0, CurrentRow.Item("Amount"))
      Date__Date = IIf(IsDBNull(CurrentRow.Item("Date")), "", CurrentRow.Item("Date"))
      Computer__String = IIf(IsDBNull(CurrentRow.Item("Computer")), "", CurrentRow.Item("Computer"))
      CheckNumber__String = IIf(IsDBNull(CurrentRow.Item("CheckNumber")), "", CurrentRow.Item("CheckNumber"))
      CheckName__String = IIf(IsDBNull(CurrentRow.Item("CheckName")), "", CurrentRow.Item("CheckName"))
      CCType__String = IIf(IsDBNull(CurrentRow.Item("CCType")), "", CurrentRow.Item("CCType"))
      CCNumber__String = IIf(IsDBNull(CurrentRow.Item("CCNumber")), "", CurrentRow.Item("CCNumber"))
      CCNumberEncrypted__String = IIf(IsDBNull(CurrentRow.Item("CCNumberEncrypted")), "", CurrentRow.Item("CCNumberEncrypted"))
      CCName__String = IIf(IsDBNull(CurrentRow.Item("CCName")), "", CurrentRow.Item("CCName"))
      CCExpMonth__String = IIf(IsDBNull(CurrentRow.Item("CCExpMonth")), "", CurrentRow.Item("CCExpMonth"))
      CCExpYear__String = IIf(IsDBNull(CurrentRow.Item("CCExpYear")), "", CurrentRow.Item("CCExpYear"))
      CCVerification__String = IIf(IsDBNull(CurrentRow.Item("CCVerification")), "", CurrentRow.Item("CCVerification"))
      CCConfirm__String = IIf(IsDBNull(CurrentRow.Item("CCConfirm")), "", CurrentRow.Item("CCConfirm"))
      QBStatus__String = IIf(IsDBNull(CurrentRow.Item("QBStatus")), "", CurrentRow.Item("QBStatus"))
      Notes__String = IIf(IsDBNull(CurrentRow.Item("Notes")), "", CurrentRow.Item("Notes"))
      User__String = IIf(IsDBNull(CurrentRow.Item("User")), "", CurrentRow.Item("User"))
      Status__String = IIf(IsDBNull(CurrentRow.Item("Status")), "", CurrentRow.Item("Status"))
      UseOnReceipt__Integer = IIf(IsDBNull(CurrentRow.Item("UseOnReceipt")), 0, CurrentRow.Item("UseOnReceipt"))
      OrigCCNumber__String = IIf(IsDBNull(CurrentRow.Item("OrigCCNumber")), "", CurrentRow.Item("OrigCCNumber"))
      CCAddress__String = IIf(IsDBNull(CurrentRow.Item("CCAddress")), "", CurrentRow.Item("CCAddress"))
      CCCity__String = IIf(IsDBNull(CurrentRow.Item("CCCity")), "", CurrentRow.Item("CCCity"))
      CCState__String = IIf(IsDBNull(CurrentRow.Item("CCState")), "", CurrentRow.Item("CCState"))
      CCZip__String = IIf(IsDBNull(CurrentRow.Item("CCZip")), "", CurrentRow.Item("CCZip"))
      AuthorizeNetReturnCode__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetReturnCode")), "", CurrentRow.Item("AuthorizeNetReturnCode"))
      AuthorizeNetReasonCode__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetReasonCode")), "", CurrentRow.Item("AuthorizeNetReasonCode"))
      AuthorizeNetReasonText__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetReasonText")), "", CurrentRow.Item("AuthorizeNetReasonText"))
      AuthorizeNetApprovalCode__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetApprovalCode")), "", CurrentRow.Item("AuthorizeNetApprovalCode"))
      AuthorizeNetAVSResultCode__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetAVSResultCode")), "", CurrentRow.Item("AuthorizeNetAVSResultCode"))
      AuthorizeNetAVSResultText__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetAVSResultText")), "", CurrentRow.Item("AuthorizeNetAVSResultText"))
      AuthorizeNetTransactionID__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetTransactionID")), "", CurrentRow.Item("AuthorizeNetTransactionID"))
      AuthorizeNetCVCCResponseCode__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetCVCCResponseCode")), "", CurrentRow.Item("AuthorizeNetCVCCResponseCode"))
      AuthorizeNetCVCCResponseText__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetCVCCResponseText")), "", CurrentRow.Item("AuthorizeNetCVCCResponseText"))
      Location__String = IIf(IsDBNull(CurrentRow.Item("Location")), "", CurrentRow.Item("Location"))
      EmailSent__String = IIf(IsDBNull(CurrentRow.Item("EmailSent")), "", CurrentRow.Item("EmailSent"))
      OriginalPaymentCategory__String = IIf(IsDBNull(CurrentRow.Item("OriginalPaymentCategory")), "", CurrentRow.Item("OriginalPaymentCategory"))
      OriginalPaymentID__Integer = IIf(IsDBNull(CurrentRow.Item("OriginalPaymentID")), 0, CurrentRow.Item("OriginalPaymentID"))
      BankName__String = IIf(IsDBNull(CurrentRow.Item("BankName")), "", CurrentRow.Item("BankName"))
      BankAccountType__String = IIf(IsDBNull(CurrentRow.Item("BankAccountType")), "", CurrentRow.Item("BankAccountType"))
      BankAccountName__String = IIf(IsDBNull(CurrentRow.Item("BankAccountName")), "", CurrentRow.Item("BankAccountName"))
      BankAccountNumber__String = IIf(IsDBNull(CurrentRow.Item("BankAccountNumber")), "", CurrentRow.Item("BankAccountNumber"))
      BankAccountNumberEncrypted__String = IIf(IsDBNull(CurrentRow.Item("BankAccountNumberEncrypted")), "", CurrentRow.Item("BankAccountNumberEncrypted"))
      BankRoutingNumber__String = IIf(IsDBNull(CurrentRow.Item("BankRoutingNumber")), "", CurrentRow.Item("BankRoutingNumber"))
      Email__String = IIf(IsDBNull(CurrentRow.Item("Email")), "", CurrentRow.Item("Email"))
      AuthorizeNetURL__String = IIf(IsDBNull(CurrentRow.Item("AuthorizeNetURL")), "", CurrentRow.Item("AuthorizeNetURL"))
      RefundPaymentID__Integer = IIf(IsDBNull(CurrentRow.Item("RefundPaymentID")), 0, CurrentRow.Item("RefundPaymentID"))
      RefundPaymentTable__String = IIf(IsDBNull(CurrentRow.Item("RefundPaymentTable")), "", CurrentRow.Item("RefundPaymentTable"))
      CheckAddress__String = IIf(IsDBNull(CurrentRow.Item("CheckAddress")), "", CurrentRow.Item("CheckAddress"))
      CheckCity__String = IIf(IsDBNull(CurrentRow.Item("CheckCity")), "", CurrentRow.Item("CheckCity"))
      CheckState__String = IIf(IsDBNull(CurrentRow.Item("CheckState")), "", CurrentRow.Item("CheckState"))
      CheckZip__String = IIf(IsDBNull(CurrentRow.Item("CheckZip")), "", CurrentRow.Item("CheckZip"))
      QBListID__String = IIf(IsDBNull(CurrentRow.Item("QBListID")), "", CurrentRow.Item("QBListID"))
      QBActivityID__Integer = IIf(IsDBNull(CurrentRow.Item("QBActivityID")), 0, CurrentRow.Item("QBActivityID"))
      AuthNETPaymentProfileID__String = IIf(IsDBNull(CurrentRow.Item("AuthNETPaymentProfileID")), "", CurrentRow.Item("AuthNETPaymentProfileID"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [RefundPayments](")
    If sInsUpdBookingID.ToString <> "" Then
      sb.Append("BookingID,")
      sbv.Append(sInsUpdBookingID & ",")
    End If
    If sInsUpdPropertyID.ToString <> "" Then
      sb.Append("PropertyID,")
      sbv.Append(sInsUpdPropertyID & ",")
    End If
    If sInsUpdHostID.ToString <> "" Then
      sb.Append("HostID,")
      sbv.Append(sInsUpdHostID & ",")
    End If
    If sInsUpdGuestID.ToString <> "" Then
      sb.Append("GuestID,")
      sbv.Append(sInsUpdGuestID & ",")
    End If
    If sInsUpdCategory.ToString <> "" Then
      sb.Append("Category,")
      sbv.Append(sInsUpdCategory & ",")
    End If
    If sInsUpdType.ToString <> "" Then
      sb.Append("Type,")
      sbv.Append(sInsUpdType & ",")
    End If
    If sInsUpdAmount.ToString <> "" Then
      sb.Append("Amount,")
      sbv.Append(sInsUpdAmount & ",")
    End If
    If sInsUpdDate.ToString <> "" Then
      sb.Append("[Date],")
      sbv.Append(sInsUpdDate & ",")
    End If
    If sInsUpdComputer.ToString <> "" Then
      sb.Append("Computer,")
      sbv.Append(sInsUpdComputer & ",")
    End If
    If sInsUpdCheckNumber.ToString <> "" Then
      sb.Append("CheckNumber,")
      sbv.Append(sInsUpdCheckNumber & ",")
    End If
    If sInsUpdCheckName.ToString <> "" Then
      sb.Append("CheckName,")
      sbv.Append(sInsUpdCheckName & ",")
    End If
    If sInsUpdCCType.ToString <> "" Then
      sb.Append("CCType,")
      sbv.Append(sInsUpdCCType & ",")
    End If
    If sInsUpdCCNumber.ToString <> "" Then
      sb.Append("CCNumber,")
      sbv.Append(sInsUpdCCNumber & ",")
    End If
    If sInsUpdCCNumberEncrypted.ToString <> "" Then
      sb.Append("CCNumberEncrypted,")
      sbv.Append(sInsUpdCCNumberEncrypted & ",")
    End If
    If sInsUpdCCName.ToString <> "" Then
      sb.Append("CCName,")
      sbv.Append(sInsUpdCCName & ",")
    End If
    If sInsUpdCCExpMonth.ToString <> "" Then
      sb.Append("CCExpMonth,")
      sbv.Append(sInsUpdCCExpMonth & ",")
    End If
    If sInsUpdCCExpYear.ToString <> "" Then
      sb.Append("CCExpYear,")
      sbv.Append(sInsUpdCCExpYear & ",")
    End If
    If sInsUpdCCVerification.ToString <> "" Then
      sb.Append("CCVerification,")
      sbv.Append(sInsUpdCCVerification & ",")
    End If
    If sInsUpdCCConfirm.ToString <> "" Then
      sb.Append("CCConfirm,")
      sbv.Append(sInsUpdCCConfirm & ",")
    End If
    If sInsUpdQBStatus.ToString <> "" Then
      sb.Append("QBStatus,")
      sbv.Append(sInsUpdQBStatus & ",")
    End If
    If sInsUpdNotes.ToString <> "" Then
      sb.Append("Notes,")
      sbv.Append(sInsUpdNotes & ",")
    End If
    If sInsUpdUser.ToString <> "" Then
      sb.Append("[User],")
      sbv.Append(sInsUpdUser & ",")
    End If
    If sInsUpdStatus.ToString <> "" Then
      sb.Append("Status,")
      sbv.Append(sInsUpdStatus & ",")
    End If
    If sInsUpdUseOnReceipt.ToString <> "" Then
      sb.Append("UseOnReceipt,")
      sbv.Append(sInsUpdUseOnReceipt & ",")
    End If
    If sInsUpdOrigCCNumber.ToString <> "" Then
      sb.Append("OrigCCNumber,")
      sbv.Append(sInsUpdOrigCCNumber & ",")
    End If
    If sInsUpdCCAddress.ToString <> "" Then
      sb.Append("CCAddress,")
      sbv.Append(sInsUpdCCAddress & ",")
    End If
    If sInsUpdCCCity.ToString <> "" Then
      sb.Append("CCCity,")
      sbv.Append(sInsUpdCCCity & ",")
    End If
    If sInsUpdCCState.ToString <> "" Then
      sb.Append("CCState,")
      sbv.Append(sInsUpdCCState & ",")
    End If
    If sInsUpdCCZip.ToString <> "" Then
      sb.Append("CCZip,")
      sbv.Append(sInsUpdCCZip & ",")
    End If
    If sInsUpdAuthorizeNetReturnCode.ToString <> "" Then
      sb.Append("AuthorizeNetReturnCode,")
      sbv.Append(sInsUpdAuthorizeNetReturnCode & ",")
    End If
    If sInsUpdAuthorizeNetReasonCode.ToString <> "" Then
      sb.Append("AuthorizeNetReasonCode,")
      sbv.Append(sInsUpdAuthorizeNetReasonCode & ",")
    End If
    If sInsUpdAuthorizeNetReasonText.ToString <> "" Then
      sb.Append("AuthorizeNetReasonText,")
      sbv.Append(sInsUpdAuthorizeNetReasonText & ",")
    End If
    If sInsUpdAuthorizeNetApprovalCode.ToString <> "" Then
      sb.Append("AuthorizeNetApprovalCode,")
      sbv.Append(sInsUpdAuthorizeNetApprovalCode & ",")
    End If
    If sInsUpdAuthorizeNetAVSResultCode.ToString <> "" Then
      sb.Append("AuthorizeNetAVSResultCode,")
      sbv.Append(sInsUpdAuthorizeNetAVSResultCode & ",")
    End If
    If sInsUpdAuthorizeNetAVSResultText.ToString <> "" Then
      sb.Append("AuthorizeNetAVSResultText,")
      sbv.Append(sInsUpdAuthorizeNetAVSResultText & ",")
    End If
    If sInsUpdAuthorizeNetTransactionID.ToString <> "" Then
      sb.Append("AuthorizeNetTransactionID,")
      sbv.Append(sInsUpdAuthorizeNetTransactionID & ",")
    End If
    If sInsUpdAuthorizeNetCVCCResponseCode.ToString <> "" Then
      sb.Append("AuthorizeNetCVCCResponseCode,")
      sbv.Append(sInsUpdAuthorizeNetCVCCResponseCode & ",")
    End If
    If sInsUpdAuthorizeNetCVCCResponseText.ToString <> "" Then
      sb.Append("AuthorizeNetCVCCResponseText,")
      sbv.Append(sInsUpdAuthorizeNetCVCCResponseText & ",")
    End If
    If sInsUpdLocation.ToString <> "" Then
      sb.Append("Location,")
      sbv.Append(sInsUpdLocation & ",")
    End If
    If sInsUpdEmailSent.ToString <> "" Then
      sb.Append("EmailSent,")
      sbv.Append(sInsUpdEmailSent & ",")
    End If
    If sInsUpdOriginalPaymentCategory.ToString <> "" Then
      sb.Append("OriginalPaymentCategory,")
      sbv.Append(sInsUpdOriginalPaymentCategory & ",")
    End If
    If sInsUpdOriginalPaymentID.ToString <> "" Then
      sb.Append("OriginalPaymentID,")
      sbv.Append(sInsUpdOriginalPaymentID & ",")
    End If
    If sInsUpdBankName.ToString <> "" Then
      sb.Append("BankName,")
      sbv.Append(sInsUpdBankName & ",")
    End If
    If sInsUpdBankAccountType.ToString <> "" Then
      sb.Append("BankAccountType,")
      sbv.Append(sInsUpdBankAccountType & ",")
    End If
    If sInsUpdBankAccountName.ToString <> "" Then
      sb.Append("BankAccountName,")
      sbv.Append(sInsUpdBankAccountName & ",")
    End If
    If sInsUpdBankAccountNumber.ToString <> "" Then
      sb.Append("BankAccountNumber,")
      sbv.Append(sInsUpdBankAccountNumber & ",")
    End If
    If sInsUpdBankAccountNumberEncrypted.ToString <> "" Then
      sb.Append("BankAccountNumberEncrypted,")
      sbv.Append(sInsUpdBankAccountNumberEncrypted & ",")
    End If
    If sInsUpdBankRoutingNumber.ToString <> "" Then
      sb.Append("BankRoutingNumber,")
      sbv.Append(sInsUpdBankRoutingNumber & ",")
    End If
    If sInsUpdEmail.ToString <> "" Then
      sb.Append("Email,")
      sbv.Append(sInsUpdEmail & ",")
    End If
    If sInsUpdAuthorizeNetURL.ToString <> "" Then
      sb.Append("AuthorizeNetURL,")
      sbv.Append(sInsUpdAuthorizeNetURL & ",")
    End If
    If sInsUpdRefundPaymentID.ToString <> "" Then
      sb.Append("RefundPaymentID,")
      sbv.Append(sInsUpdRefundPaymentID & ",")
    End If
    If sInsUpdRefundPaymentTable.ToString <> "" Then
      sb.Append("RefundPaymentTable,")
      sbv.Append(sInsUpdRefundPaymentTable & ",")
    End If
    If sInsUpdCheckAddress.ToString <> "" Then
      sb.Append("CheckAddress,")
      sbv.Append(sInsUpdCheckAddress & ",")
    End If
    If sInsUpdCheckCity.ToString <> "" Then
      sb.Append("CheckCity,")
      sbv.Append(sInsUpdCheckCity & ",")
    End If
    If sInsUpdCheckState.ToString <> "" Then
      sb.Append("CheckState,")
      sbv.Append(sInsUpdCheckState & ",")
    End If
    If sInsUpdCheckZip.ToString <> "" Then
      sb.Append("CheckZip,")
      sbv.Append(sInsUpdCheckZip & ",")
    End If
    If sInsUpdQBListID.ToString <> "" Then
      sb.Append("QBListID,")
      sbv.Append(sInsUpdQBListID & ",")
    End If
    If sInsUpdQBActivityID.ToString <> "" Then
      sb.Append("QBActivityID,")
      sbv.Append(sInsUpdQBActivityID & ",")
    End If
    If sInsUpdAuthNETPaymentProfileID.ToString <> "" Then
      sb.Append("AuthNETPaymentProfileID,")
      sbv.Append(sInsUpdAuthNETPaymentProfileID & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(PaymentID) from [RefundPayments]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    PaymentID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [RefundPayments] Set ")
    If sInsUpdBookingID.ToString <> "" Then sb.Append("BookingID=" & sInsUpdBookingID & ",")
    If sInsUpdPropertyID.ToString <> "" Then sb.Append("PropertyID=" & sInsUpdPropertyID & ",")
    If sInsUpdHostID.ToString <> "" Then sb.Append("HostID=" & sInsUpdHostID & ",")
    If sInsUpdGuestID.ToString <> "" Then sb.Append("GuestID=" & sInsUpdGuestID & ",")
    If sInsUpdCategory.ToString <> "" Then sb.Append("Category=" & sInsUpdCategory & ",")
    If sInsUpdType.ToString <> "" Then sb.Append("Type=" & sInsUpdType & ",")
    If sInsUpdAmount.ToString <> "" Then sb.Append("Amount=" & sInsUpdAmount & ",")
    If sInsUpdDate.ToString <> "" Then sb.Append("[Date]=" & sInsUpdDate & ",")
    If sInsUpdComputer.ToString <> "" Then sb.Append("Computer=" & sInsUpdComputer & ",")
    If sInsUpdCheckNumber.ToString <> "" Then sb.Append("CheckNumber=" & sInsUpdCheckNumber & ",")
    If sInsUpdCheckName.ToString <> "" Then sb.Append("CheckName=" & sInsUpdCheckName & ",")
    If sInsUpdCCType.ToString <> "" Then sb.Append("CCType=" & sInsUpdCCType & ",")
    If sInsUpdCCNumber.ToString <> "" Then sb.Append("CCNumber=" & sInsUpdCCNumber & ",")
    If sInsUpdCCNumberEncrypted.ToString <> "" Then sb.Append("CCNumberEncrypted=" & sInsUpdCCNumberEncrypted & ",")
    If sInsUpdCCName.ToString <> "" Then sb.Append("CCName=" & sInsUpdCCName & ",")
    If sInsUpdCCExpMonth.ToString <> "" Then sb.Append("CCExpMonth=" & sInsUpdCCExpMonth & ",")
    If sInsUpdCCExpYear.ToString <> "" Then sb.Append("CCExpYear=" & sInsUpdCCExpYear & ",")
    If sInsUpdCCVerification.ToString <> "" Then sb.Append("CCVerification=" & sInsUpdCCVerification & ",")
    If sInsUpdCCConfirm.ToString <> "" Then sb.Append("CCConfirm=" & sInsUpdCCConfirm & ",")
    If sInsUpdQBStatus.ToString <> "" Then sb.Append("QBStatus=" & sInsUpdQBStatus & ",")
    If sInsUpdNotes.ToString <> "" Then sb.Append("Notes=" & sInsUpdNotes & ",")
    If sInsUpdUser.ToString <> "" Then sb.Append("[User]=" & sInsUpdUser & ",")
    If sInsUpdStatus.ToString <> "" Then sb.Append("Status=" & sInsUpdStatus & ",")
    If sInsUpdUseOnReceipt.ToString <> "" Then sb.Append("UseOnReceipt=" & sInsUpdUseOnReceipt & ",")
    If sInsUpdOrigCCNumber.ToString <> "" Then sb.Append("OrigCCNumber=" & sInsUpdOrigCCNumber & ",")
    If sInsUpdCCAddress.ToString <> "" Then sb.Append("CCAddress=" & sInsUpdCCAddress & ",")
    If sInsUpdCCCity.ToString <> "" Then sb.Append("CCCity=" & sInsUpdCCCity & ",")
    If sInsUpdCCState.ToString <> "" Then sb.Append("CCState=" & sInsUpdCCState & ",")
    If sInsUpdCCZip.ToString <> "" Then sb.Append("CCZip=" & sInsUpdCCZip & ",")
    If sInsUpdAuthorizeNetReturnCode.ToString <> "" Then sb.Append("AuthorizeNetReturnCode=" & sInsUpdAuthorizeNetReturnCode & ",")
    If sInsUpdAuthorizeNetReasonCode.ToString <> "" Then sb.Append("AuthorizeNetReasonCode=" & sInsUpdAuthorizeNetReasonCode & ",")
    If sInsUpdAuthorizeNetReasonText.ToString <> "" Then sb.Append("AuthorizeNetReasonText=" & sInsUpdAuthorizeNetReasonText & ",")
    If sInsUpdAuthorizeNetApprovalCode.ToString <> "" Then sb.Append("AuthorizeNetApprovalCode=" & sInsUpdAuthorizeNetApprovalCode & ",")
    If sInsUpdAuthorizeNetAVSResultCode.ToString <> "" Then sb.Append("AuthorizeNetAVSResultCode=" & sInsUpdAuthorizeNetAVSResultCode & ",")
    If sInsUpdAuthorizeNetAVSResultText.ToString <> "" Then sb.Append("AuthorizeNetAVSResultText=" & sInsUpdAuthorizeNetAVSResultText & ",")
    If sInsUpdAuthorizeNetTransactionID.ToString <> "" Then sb.Append("AuthorizeNetTransactionID=" & sInsUpdAuthorizeNetTransactionID & ",")
    If sInsUpdAuthorizeNetCVCCResponseCode.ToString <> "" Then sb.Append("AuthorizeNetCVCCResponseCode=" & sInsUpdAuthorizeNetCVCCResponseCode & ",")
    If sInsUpdAuthorizeNetCVCCResponseText.ToString <> "" Then sb.Append("AuthorizeNetCVCCResponseText=" & sInsUpdAuthorizeNetCVCCResponseText & ",")
    If sInsUpdLocation.ToString <> "" Then sb.Append("Location=" & sInsUpdLocation & ",")
    If sInsUpdEmailSent.ToString <> "" Then sb.Append("EmailSent=" & sInsUpdEmailSent & ",")
    If sInsUpdOriginalPaymentCategory.ToString <> "" Then sb.Append("OriginalPaymentCategory=" & sInsUpdOriginalPaymentCategory & ",")
    If sInsUpdOriginalPaymentID.ToString <> "" Then sb.Append("OriginalPaymentID=" & sInsUpdOriginalPaymentID & ",")
    If sInsUpdBankName.ToString <> "" Then sb.Append("BankName=" & sInsUpdBankName & ",")
    If sInsUpdBankAccountType.ToString <> "" Then sb.Append("BankAccountType=" & sInsUpdBankAccountType & ",")
    If sInsUpdBankAccountName.ToString <> "" Then sb.Append("BankAccountName=" & sInsUpdBankAccountName & ",")
    If sInsUpdBankAccountNumber.ToString <> "" Then sb.Append("BankAccountNumber=" & sInsUpdBankAccountNumber & ",")
    If sInsUpdBankAccountNumberEncrypted.ToString <> "" Then sb.Append("BankAccountNumberEncrypted=" & sInsUpdBankAccountNumberEncrypted & ",")
    If sInsUpdBankRoutingNumber.ToString <> "" Then sb.Append("BankRoutingNumber=" & sInsUpdBankRoutingNumber & ",")
    If sInsUpdEmail.ToString <> "" Then sb.Append("Email=" & sInsUpdEmail & ",")
    If sInsUpdAuthorizeNetURL.ToString <> "" Then sb.Append("AuthorizeNetURL=" & sInsUpdAuthorizeNetURL & ",")
    If sInsUpdRefundPaymentID.ToString <> "" Then sb.Append("RefundPaymentID=" & sInsUpdRefundPaymentID & ",")
    If sInsUpdRefundPaymentTable.ToString <> "" Then sb.Append("RefundPaymentTable=" & sInsUpdRefundPaymentTable & ",")
    If sInsUpdCheckAddress.ToString <> "" Then sb.Append("CheckAddress=" & sInsUpdCheckAddress & ",")
    If sInsUpdCheckCity.ToString <> "" Then sb.Append("CheckCity=" & sInsUpdCheckCity & ",")
    If sInsUpdCheckState.ToString <> "" Then sb.Append("CheckState=" & sInsUpdCheckState & ",")
    If sInsUpdCheckZip.ToString <> "" Then sb.Append("CheckZip=" & sInsUpdCheckZip & ",")
    If sInsUpdQBListID.ToString <> "" Then sb.Append("QBListID=" & sInsUpdQBListID & ",")
    If sInsUpdQBActivityID.ToString <> "" Then sb.Append("QBActivityID=" & sInsUpdQBActivityID & ",")
    If sInsUpdAuthNETPaymentProfileID.ToString <> "" Then sb.Append("AuthNETPaymentProfileID=" & sInsUpdAuthNETPaymentProfileID & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where PaymentID=" & sInsUpdPaymentID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [RefundPayments] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("PaymentID=" & sInsUpdPaymentID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TableEmail

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iidEmail As Int32
  Private sInsUpdidEmail As String
  Property idEmail_PK__Integer() As Int32
    Get
      Return iidEmail
    End Get
    Set(ByVal Value As Int32)
      iidEmail = Value
      sInsUpdidEmail = oUtil.FixParam(iidEmail, True)
    End Set
  End Property

  Private ibEmailFormatIsHTML As Int32
  Private sInsUpdbEmailFormatIsHTML As String
  Property bEmailFormatIsHTML_RQ__Integer() As Int32
    Get
      Return ibEmailFormatIsHTML
    End Get
    Set(ByVal Value As Int32)
      ibEmailFormatIsHTML = Value
      sInsUpdbEmailFormatIsHTML = oUtil.FixParam(ibEmailFormatIsHTML, False)
    End Set
  End Property

  Private scFromAddress As String
  Private sInsUpdcFromAddress As String
  Property cFromAddress_RQ__String() As String
    Get
      Return scFromAddress
    End Get
    Set(ByVal Value As String)
      scFromAddress = Value
      sInsUpdcFromAddress = oUtil.FixParam(scFromAddress, False)
    End Set
  End Property

  Private scFromName As String
  Private sInsUpdcFromName As String
  Property cFromName__String() As String
    Get
      Return scFromName
    End Get
    Set(ByVal Value As String)
      scFromName = Value
      sInsUpdcFromName = oUtil.FixParam(scFromName, True)
    End Set
  End Property

  Private scToAddress As String
  Private sInsUpdcToAddress As String
  Property cToAddress_RQ__String() As String
    Get
      Return scToAddress
    End Get
    Set(ByVal Value As String)
      scToAddress = Value
      sInsUpdcToAddress = oUtil.FixParam(scToAddress, False)
    End Set
  End Property

  Private scToName As String
  Private sInsUpdcToName As String
  Property cToName__String() As String
    Get
      Return scToName
    End Get
    Set(ByVal Value As String)
      scToName = Value
      sInsUpdcToName = oUtil.FixParam(scToName, True)
    End Set
  End Property

  Private scSubject As String
  Private sInsUpdcSubject As String
  Property cSubject_RQ__String() As String
    Get
      Return scSubject
    End Get
    Set(ByVal Value As String)
      scSubject = Value
      sInsUpdcSubject = oUtil.FixParam(scSubject, False)
    End Set
  End Property

  Private scBody As String
  Private sInsUpdcBody As String
  Property cBody_RQ__String() As String
    Get
      Return scBody
    End Get
    Set(ByVal Value As String)
      scBody = Value
      sInsUpdcBody = oUtil.FixParam(scBody, False)
    End Set
  End Property

  Private scAttachmentFile As String
  Private sInsUpdcAttachmentFile As String
  Property cAttachmentFile__String() As String
    Get
      Return scAttachmentFile
    End Get
    Set(ByVal Value As String)
      scAttachmentFile = Value
      sInsUpdcAttachmentFile = oUtil.FixParam(scAttachmentFile, True)
    End Set
  End Property

  Private scStatus As String
  Private sInsUpdcStatus As String
  Property cStatus_RQ__String() As String
    Get
      Return scStatus
    End Get
    Set(ByVal Value As String)
      scStatus = Value
      sInsUpdcStatus = oUtil.FixParam(scStatus, False)
    End Set
  End Property

  Private sdtAddDate As String
  Private sInsUpddtAddDate As String
  Property dtAddDate__Date() As String
    Get
      Return sdtAddDate
    End Get
    Set(ByVal Value As String)
      sdtAddDate = Value
      sInsUpddtAddDate = oUtil.FixParam(sdtAddDate, True)
    End Set
  End Property

  Private sdtSendDate As String
  Private sInsUpddtSendDate As String
  Property dtSendDate__Date() As String
    Get
      Return sdtSendDate
    End Get
    Set(ByVal Value As String)
      sdtSendDate = Value
      sInsUpddtSendDate = oUtil.FixParam(sdtSendDate, True)
    End Set
  End Property

  Private iiPriority As Int32
  Private sInsUpdiPriority As String
  Property iPriority__Integer() As Int32
    Get
      Return iiPriority
    End Get
    Set(ByVal Value As Int32)
      iiPriority = Value
      sInsUpdiPriority = oUtil.FixParam(iiPriority, True)
    End Set
  End Property

  Private scSubmittedBy As String
  Private sInsUpdcSubmittedBy As String
  Property cSubmittedBy__String() As String
    Get
      Return scSubmittedBy
    End Get
    Set(ByVal Value As String)
      scSubmittedBy = Value
      sInsUpdcSubmittedBy = oUtil.FixParam(scSubmittedBy, True)
    End Set
  End Property

  Public Sub Clear()
    iidEmail = 0
    sInsUpdidEmail = ""
    ibEmailFormatIsHTML = 0
    sInsUpdbEmailFormatIsHTML = ""
    scFromAddress = ""
    sInsUpdcFromAddress = ""
    scFromName = ""
    sInsUpdcFromName = ""
    scToAddress = ""
    sInsUpdcToAddress = ""
    scToName = ""
    sInsUpdcToName = ""
    scSubject = ""
    sInsUpdcSubject = ""
    scBody = ""
    sInsUpdcBody = ""
    scAttachmentFile = ""
    sInsUpdcAttachmentFile = ""
    scStatus = ""
    sInsUpdcStatus = ""
    sdtAddDate = ""
    sInsUpddtAddDate = ""
    sdtSendDate = ""
    sInsUpddtSendDate = ""
    iiPriority = 0
    sInsUpdiPriority = ""
    scSubmittedBy = ""
    sInsUpdcSubmittedBy = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdidEmail.ToString = "'-12345'" Then sb.Append("idEmail,")
        If sInsUpdbEmailFormatIsHTML.ToString = "'-12345'" Then sb.Append("bEmailFormatIsHTML,")
        If sInsUpdcFromAddress.ToString = "'Select'" Then sb.Append("cFromAddress,")
        If sInsUpdcFromName.ToString = "'Select'" Then sb.Append("cFromName,")
        If sInsUpdcToAddress.ToString = "'Select'" Then sb.Append("cToAddress,")
        If sInsUpdcToName.ToString = "'Select'" Then sb.Append("cToName,")
        If sInsUpdcSubject.ToString = "'Select'" Then sb.Append("cSubject,")
        If sInsUpdcBody.ToString = "'Select'" Then sb.Append("cBody,")
        If sInsUpdcAttachmentFile.ToString = "'Select'" Then sb.Append("cAttachmentFile,")
        If sInsUpdcStatus.ToString = "'Select'" Then sb.Append("cStatus,")
        If sInsUpddtAddDate.ToString = "'Select'" Then sb.Append("dtAddDate,")
        If sInsUpddtSendDate.ToString = "'Select'" Then sb.Append("dtSendDate,")
        If sInsUpdiPriority.ToString = "'-12345'" Then sb.Append("iPriority,")
        If sInsUpdcSubmittedBy.ToString = "'Select'" Then sb.Append("cSubmittedBy,")
      Else
        sb.Append("idEmail,")
        sb.Append("bEmailFormatIsHTML,")
        sb.Append("cFromAddress,")
        sb.Append("cFromName,")
        sb.Append("cToAddress,")
        sb.Append("cToName,")
        sb.Append("cSubject,")
        sb.Append("cBody,")
        sb.Append("cAttachmentFile,")
        sb.Append("cStatus,")
        sb.Append("dtAddDate,")
        sb.Append("dtSendDate,")
        sb.Append("iPriority,")
        sb.Append("cSubmittedBy,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [Email]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdidEmail.ToString <> "") And (sInsUpdidEmail <> "'-12345'") Then sbw.Append("idEmail=" & sInsUpdidEmail & " and ")
      If (sInsUpdbEmailFormatIsHTML.ToString <> "") And (sInsUpdbEmailFormatIsHTML <> "'-12345'") Then sbw.Append("bEmailFormatIsHTML=" & sInsUpdbEmailFormatIsHTML & " and ")
      If (sInsUpdcFromAddress.ToString <> "") And (sInsUpdcFromAddress <> "'Select'") Then sbw.Append("cFromAddress=" & sInsUpdcFromAddress & " and ")
      If (sInsUpdcFromName.ToString <> "") And (sInsUpdcFromName <> "'Select'") Then sbw.Append("cFromName=" & sInsUpdcFromName & " and ")
      If (sInsUpdcToAddress.ToString <> "") And (sInsUpdcToAddress <> "'Select'") Then sbw.Append("cToAddress=" & sInsUpdcToAddress & " and ")
      If (sInsUpdcToName.ToString <> "") And (sInsUpdcToName <> "'Select'") Then sbw.Append("cToName=" & sInsUpdcToName & " and ")
      If (sInsUpdcSubject.ToString <> "") And (sInsUpdcSubject <> "'Select'") Then sbw.Append("cSubject=" & sInsUpdcSubject & " and ")
      If (sInsUpdcBody.ToString <> "") And (sInsUpdcBody <> "'Select'") Then sbw.Append("cBody=" & sInsUpdcBody & " and ")
      If (sInsUpdcAttachmentFile.ToString <> "") And (sInsUpdcAttachmentFile <> "'Select'") Then sbw.Append("cAttachmentFile=" & sInsUpdcAttachmentFile & " and ")
      If (sInsUpdcStatus.ToString <> "") And (sInsUpdcStatus <> "'Select'") Then sbw.Append("cStatus=" & sInsUpdcStatus & " and ")
      If (sInsUpddtAddDate.ToString <> "") And (sInsUpddtAddDate <> "'Select'") Then sbw.Append("dtAddDate=" & sInsUpddtAddDate & " and ")
      If (sInsUpddtSendDate.ToString <> "") And (sInsUpddtSendDate <> "'Select'") Then sbw.Append("dtSendDate=" & sInsUpddtSendDate & " and ")
      If (sInsUpdiPriority.ToString <> "") And (sInsUpdiPriority <> "'-12345'") Then sbw.Append("iPriority=" & sInsUpdiPriority & " and ")
      If (sInsUpdcSubmittedBy.ToString <> "") And (sInsUpdcSubmittedBy <> "'Select'") Then sbw.Append("cSubmittedBy=" & sInsUpdcSubmittedBy & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      idEmail_PK__Integer = IIf(IsDBNull(CurrentRow.Item("idEmail")), 0, CurrentRow.Item("idEmail").ToString)
      bEmailFormatIsHTML_RQ__Integer = IIf(IsDBNull(CurrentRow.Item("bEmailFormatIsHTML")), 0, CurrentRow.Item("bEmailFormatIsHTML"))
      cFromAddress_RQ__String = IIf(IsDBNull(CurrentRow.Item("cFromAddress")), "", CurrentRow.Item("cFromAddress"))
      cFromName__String = IIf(IsDBNull(CurrentRow.Item("cFromName")), "", CurrentRow.Item("cFromName"))
      cToAddress_RQ__String = IIf(IsDBNull(CurrentRow.Item("cToAddress")), "", CurrentRow.Item("cToAddress"))
      cToName__String = IIf(IsDBNull(CurrentRow.Item("cToName")), "", CurrentRow.Item("cToName"))
      cSubject_RQ__String = IIf(IsDBNull(CurrentRow.Item("cSubject")), "", CurrentRow.Item("cSubject"))
      cBody_RQ__String = IIf(IsDBNull(CurrentRow.Item("cBody")), "", CurrentRow.Item("cBody"))
      cAttachmentFile__String = IIf(IsDBNull(CurrentRow.Item("cAttachmentFile")), "", CurrentRow.Item("cAttachmentFile"))
      cStatus_RQ__String = IIf(IsDBNull(CurrentRow.Item("cStatus")), "", CurrentRow.Item("cStatus"))
      dtAddDate__Date = IIf(IsDBNull(CurrentRow.Item("dtAddDate")), "", CurrentRow.Item("dtAddDate"))
      dtSendDate__Date = IIf(IsDBNull(CurrentRow.Item("dtSendDate")), "", CurrentRow.Item("dtSendDate"))
      iPriority__Integer = IIf(IsDBNull(CurrentRow.Item("iPriority")), 0, CurrentRow.Item("iPriority"))
      cSubmittedBy__String = IIf(IsDBNull(CurrentRow.Item("cSubmittedBy")), "", CurrentRow.Item("cSubmittedBy"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [Email](")
    If sInsUpdbEmailFormatIsHTML.ToString <> "" Then
      sb.Append("bEmailFormatIsHTML,")
      sbv.Append(sInsUpdbEmailFormatIsHTML & ",")
    End If
    If sInsUpdcFromAddress.ToString <> "" Then
      sb.Append("cFromAddress,")
      sbv.Append(sInsUpdcFromAddress & ",")
    End If
    If sInsUpdcFromName.ToString <> "" Then
      sb.Append("cFromName,")
      sbv.Append(sInsUpdcFromName & ",")
    End If
    If sInsUpdcToAddress.ToString <> "" Then
      sb.Append("cToAddress,")
      sbv.Append(sInsUpdcToAddress & ",")
    End If
    If sInsUpdcToName.ToString <> "" Then
      sb.Append("cToName,")
      sbv.Append(sInsUpdcToName & ",")
    End If
    If sInsUpdcSubject.ToString <> "" Then
      sb.Append("cSubject,")
      sbv.Append(sInsUpdcSubject & ",")
    End If
    If sInsUpdcBody.ToString <> "" Then
      sb.Append("cBody,")
      sbv.Append(sInsUpdcBody & ",")
    End If
    If sInsUpdcAttachmentFile.ToString <> "" Then
      sb.Append("cAttachmentFile,")
      sbv.Append(sInsUpdcAttachmentFile & ",")
    End If
    If sInsUpdcStatus.ToString <> "" Then
      sb.Append("cStatus,")
      sbv.Append(sInsUpdcStatus & ",")
    End If
    If sInsUpddtAddDate.ToString <> "" Then
      sb.Append("dtAddDate,")
      sbv.Append(sInsUpddtAddDate & ",")
    End If
    If sInsUpddtSendDate.ToString <> "" Then
      sb.Append("dtSendDate,")
      sbv.Append(sInsUpddtSendDate & ",")
    End If
    If sInsUpdiPriority.ToString <> "" Then
      sb.Append("iPriority,")
      sbv.Append(sInsUpdiPriority & ",")
    End If
    If sInsUpdcSubmittedBy.ToString <> "" Then
      sb.Append("cSubmittedBy,")
      sbv.Append(sInsUpdcSubmittedBy & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(idEmail) from [Email]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    idEmail_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [Email] Set ")
    If sInsUpdbEmailFormatIsHTML.ToString <> "" Then sb.Append("bEmailFormatIsHTML=" & sInsUpdbEmailFormatIsHTML & ",")
    If sInsUpdcFromAddress.ToString <> "" Then sb.Append("cFromAddress=" & sInsUpdcFromAddress & ",")
    If sInsUpdcFromName.ToString <> "" Then sb.Append("cFromName=" & sInsUpdcFromName & ",")
    If sInsUpdcToAddress.ToString <> "" Then sb.Append("cToAddress=" & sInsUpdcToAddress & ",")
    If sInsUpdcToName.ToString <> "" Then sb.Append("cToName=" & sInsUpdcToName & ",")
    If sInsUpdcSubject.ToString <> "" Then sb.Append("cSubject=" & sInsUpdcSubject & ",")
    If sInsUpdcBody.ToString <> "" Then sb.Append("cBody=" & sInsUpdcBody & ",")
    If sInsUpdcAttachmentFile.ToString <> "" Then sb.Append("cAttachmentFile=" & sInsUpdcAttachmentFile & ",")
    If sInsUpdcStatus.ToString <> "" Then sb.Append("cStatus=" & sInsUpdcStatus & ",")
    If sInsUpddtAddDate.ToString <> "" Then sb.Append("dtAddDate=" & sInsUpddtAddDate & ",")
    If sInsUpddtSendDate.ToString <> "" Then sb.Append("dtSendDate=" & sInsUpddtSendDate & ",")
    If sInsUpdiPriority.ToString <> "" Then sb.Append("iPriority=" & sInsUpdiPriority & ",")
    If sInsUpdcSubmittedBy.ToString <> "" Then sb.Append("cSubmittedBy=" & sInsUpdcSubmittedBy & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where idEmail=" & sInsUpdidEmail
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [Email] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("idEmail=" & sInsUpdidEmail)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TablePropertyCategoryURLs

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iProperty_ID As Int32
  Private sInsUpdProperty_ID As String
  Property Property_ID_PK__Integer() As Int32
    Get
      Return iProperty_ID
    End Get
    Set(ByVal Value As Int32)
      iProperty_ID = Value
      sInsUpdProperty_ID = oUtil.FixParam(iProperty_ID, True)
    End Set
  End Property

  Private sURL As String
  Private sInsUpdURL As String
  Property URL__String() As String
    Get
      Return sURL
    End Get
    Set(ByVal Value As String)
      sURL = Value
      sInsUpdURL = oUtil.FixParam(sURL, True)
    End Set
  End Property

  Private sCategory_ID As String
  Private sInsUpdCategory_ID As String
  Property Category_ID__String() As String
    Get
      Return sCategory_ID
    End Get
    Set(ByVal Value As String)
      sCategory_ID = Value
      sInsUpdCategory_ID = oUtil.FixParam(sCategory_ID, True)
    End Set
  End Property

  Public Sub Clear()
    iProperty_ID = 0
    sInsUpdProperty_ID = ""
    sURL = ""
    sInsUpdURL = ""
    sCategory_ID = ""
    sInsUpdCategory_ID = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdProperty_ID.ToString = "'-12345'" Then sb.Append("Property_ID,")
        If sInsUpdURL.ToString = "'Select'" Then sb.Append("URL,")
        If sInsUpdCategory_ID.ToString = "'Select'" Then sb.Append("Category_ID,")
      Else
        sb.Append("Property_ID,")
        sb.Append("URL,")
        sb.Append("Category_ID,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [PropertyCategoryURLs]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdProperty_ID.ToString <> "") And (sInsUpdProperty_ID <> "'-12345'") Then sbw.Append("Property_ID=" & sInsUpdProperty_ID & " and ")
      If (sInsUpdURL.ToString <> "") And (sInsUpdURL <> "'Select'") Then sbw.Append("URL=" & sInsUpdURL & " and ")
      If (sInsUpdCategory_ID.ToString <> "") And (sInsUpdCategory_ID <> "'Select'") Then sbw.Append("Category_ID=" & sInsUpdCategory_ID & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      Property_ID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("Property_ID")), 0, CurrentRow.Item("Property_ID").ToString)
      URL__String = IIf(IsDBNull(CurrentRow.Item("URL")), "", CurrentRow.Item("URL"))
      Category_ID__String = IIf(IsDBNull(CurrentRow.Item("Category_ID")), "", CurrentRow.Item("Category_ID"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [PropertyCategoryURLs](")
    If sInsUpdURL.ToString <> "" Then
      sb.Append("URL,")
      sbv.Append(sInsUpdURL & ",")
    End If
    If sInsUpdCategory_ID.ToString <> "" Then
      sb.Append("Category_ID,")
      sbv.Append(sInsUpdCategory_ID & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(Property_ID) from [PropertyCategoryURLs]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    Property_ID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [PropertyCategoryURLs] Set ")
    If sInsUpdURL.ToString <> "" Then sb.Append("URL=" & sInsUpdURL & ",")
    If sInsUpdCategory_ID.ToString <> "" Then sb.Append("Category_ID=" & sInsUpdCategory_ID & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where Property_ID=" & sInsUpdProperty_ID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [PropertyCategoryURLs] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("Property_ID=" & sInsUpdProperty_ID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TableReports

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iRpt_ID As Int32
  Private sInsUpdRpt_ID As String
  Property Rpt_ID_PK__Integer() As Int32
    Get
      Return iRpt_ID
    End Get
    Set(ByVal Value As Int32)
      iRpt_ID = Value
      sInsUpdRpt_ID = oUtil.FixParam(iRpt_ID, True)
    End Set
  End Property

  Private sReportName As String
  Private sInsUpdReportName As String
  Property ReportName__String() As String
    Get
      Return sReportName
    End Get
    Set(ByVal Value As String)
      sReportName = Value
      sInsUpdReportName = oUtil.FixParam(sReportName, True)
    End Set
  End Property

  Private sReportDescription As String
  Private sInsUpdReportDescription As String
  Property ReportDescription__String() As String
    Get
      Return sReportDescription
    End Get
    Set(ByVal Value As String)
      sReportDescription = Value
      sInsUpdReportDescription = oUtil.FixParam(sReportDescription, True)
    End Set
  End Property

  Private sViewName As String
  Private sInsUpdViewName As String
  Property ViewName__String() As String
    Get
      Return sViewName
    End Get
    Set(ByVal Value As String)
      sViewName = Value
      sInsUpdViewName = oUtil.FixParam(sViewName, True)
    End Set
  End Property

  Private sUseSelectionFormula As String
  Private sInsUpdUseSelectionFormula As String
  Property UseSelectionFormula__String() As String
    Get
      Return sUseSelectionFormula
    End Get
    Set(ByVal Value As String)
      sUseSelectionFormula = Value
      sInsUpdUseSelectionFormula = oUtil.FixParam(sUseSelectionFormula, True)
    End Set
  End Property

  Public Sub Clear()
    iRpt_ID = 0
    sInsUpdRpt_ID = ""
    sReportName = ""
    sInsUpdReportName = ""
    sReportDescription = ""
    sInsUpdReportDescription = ""
    sViewName = ""
    sInsUpdViewName = ""
    sUseSelectionFormula = ""
    sInsUpdUseSelectionFormula = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdRpt_ID.ToString = "'-12345'" Then sb.Append("Rpt_ID,")
        If sInsUpdReportName.ToString = "'Select'" Then sb.Append("ReportName,")
        If sInsUpdReportDescription.ToString = "'Select'" Then sb.Append("ReportDescription,")
        If sInsUpdViewName.ToString = "'Select'" Then sb.Append("ViewName,")
        If sInsUpdUseSelectionFormula.ToString = "'Select'" Then sb.Append("UseSelectionFormula,")
      Else
        sb.Append("Rpt_ID,")
        sb.Append("ReportName,")
        sb.Append("ReportDescription,")
        sb.Append("ViewName,")
        sb.Append("UseSelectionFormula,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [Reports]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdRpt_ID.ToString <> "") And (sInsUpdRpt_ID <> "'-12345'") Then sbw.Append("Rpt_ID=" & sInsUpdRpt_ID & " and ")
      If (sInsUpdReportName.ToString <> "") And (sInsUpdReportName <> "'Select'") Then sbw.Append("ReportName=" & sInsUpdReportName & " and ")
      If (sInsUpdReportDescription.ToString <> "") And (sInsUpdReportDescription <> "'Select'") Then sbw.Append("ReportDescription=" & sInsUpdReportDescription & " and ")
      If (sInsUpdViewName.ToString <> "") And (sInsUpdViewName <> "'Select'") Then sbw.Append("ViewName=" & sInsUpdViewName & " and ")
      If (sInsUpdUseSelectionFormula.ToString <> "") And (sInsUpdUseSelectionFormula <> "'Select'") Then sbw.Append("UseSelectionFormula=" & sInsUpdUseSelectionFormula & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      Rpt_ID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("Rpt_ID")), 0, CurrentRow.Item("Rpt_ID").ToString)
      ReportName__String = IIf(IsDBNull(CurrentRow.Item("ReportName")), "", CurrentRow.Item("ReportName"))
      ReportDescription__String = IIf(IsDBNull(CurrentRow.Item("ReportDescription")), "", CurrentRow.Item("ReportDescription"))
      ViewName__String = IIf(IsDBNull(CurrentRow.Item("ViewName")), "", CurrentRow.Item("ViewName"))
      UseSelectionFormula__String = IIf(IsDBNull(CurrentRow.Item("UseSelectionFormula")), "", CurrentRow.Item("UseSelectionFormula"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [Reports](")
    If sInsUpdReportName.ToString <> "" Then
      sb.Append("ReportName,")
      sbv.Append(sInsUpdReportName & ",")
    End If
    If sInsUpdReportDescription.ToString <> "" Then
      sb.Append("ReportDescription,")
      sbv.Append(sInsUpdReportDescription & ",")
    End If
    If sInsUpdViewName.ToString <> "" Then
      sb.Append("ViewName,")
      sbv.Append(sInsUpdViewName & ",")
    End If
    If sInsUpdUseSelectionFormula.ToString <> "" Then
      sb.Append("UseSelectionFormula,")
      sbv.Append(sInsUpdUseSelectionFormula & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(Rpt_ID) from [Reports]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    Rpt_ID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [Reports] Set ")
    If sInsUpdReportName.ToString <> "" Then sb.Append("ReportName=" & sInsUpdReportName & ",")
    If sInsUpdReportDescription.ToString <> "" Then sb.Append("ReportDescription=" & sInsUpdReportDescription & ",")
    If sInsUpdViewName.ToString <> "" Then sb.Append("ViewName=" & sInsUpdViewName & ",")
    If sInsUpdUseSelectionFormula.ToString <> "" Then sb.Append("UseSelectionFormula=" & sInsUpdUseSelectionFormula & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where Rpt_ID=" & sInsUpdRpt_ID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [Reports] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("Rpt_ID=" & sInsUpdRpt_ID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TablePropertiesGroups

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iID As Int32
  Private sInsUpdID As String
  Property ID_PK__Integer() As Int32
    Get
      Return iID
    End Get
    Set(ByVal Value As Int32)
      iID = Value
      sInsUpdID = oUtil.FixParam(iID, True)
    End Set
  End Property

  Private iProperty_ID As Int32
  Private sInsUpdProperty_ID As String
  Property Property_ID_RQ__Integer() As Int32
    Get
      Return iProperty_ID
    End Get
    Set(ByVal Value As Int32)
      iProperty_ID = Value
      sInsUpdProperty_ID = oUtil.FixParam(iProperty_ID, False)
    End Set
  End Property

  Private iGroupID As Int32
  Private sInsUpdGroupID As String
  Property GroupID_RQ__Integer() As Int32
    Get
      Return iGroupID
    End Get
    Set(ByVal Value As Int32)
      iGroupID = Value
      sInsUpdGroupID = oUtil.FixParam(iGroupID, False)
    End Set
  End Property

  Private iSequenceNumber As Int32
  Private sInsUpdSequenceNumber As String
  Property SequenceNumber_RQ__Integer() As Int32
    Get
      Return iSequenceNumber
    End Get
    Set(ByVal Value As Int32)
      iSequenceNumber = Value
      sInsUpdSequenceNumber = oUtil.FixParam(iSequenceNumber, False)
    End Set
  End Property

  Public Sub Clear()
    iID = 0
    sInsUpdID = ""
    iProperty_ID = 0
    sInsUpdProperty_ID = ""
    iGroupID = 0
    sInsUpdGroupID = ""
    iSequenceNumber = 0
    sInsUpdSequenceNumber = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdID.ToString = "'-12345'" Then sb.Append("ID,")
        If sInsUpdProperty_ID.ToString = "'-12345'" Then sb.Append("Property_ID,")
        If sInsUpdGroupID.ToString = "'-12345'" Then sb.Append("GroupID,")
        If sInsUpdSequenceNumber.ToString = "'-12345'" Then sb.Append("SequenceNumber,")
      Else
        sb.Append("ID,")
        sb.Append("Property_ID,")
        sb.Append("GroupID,")
        sb.Append("SequenceNumber,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [PropertiesGroups]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdID.ToString <> "") And (sInsUpdID <> "'-12345'") Then sbw.Append("ID=" & sInsUpdID & " and ")
      If (sInsUpdProperty_ID.ToString <> "") And (sInsUpdProperty_ID <> "'-12345'") Then sbw.Append("Property_ID=" & sInsUpdProperty_ID & " and ")
      If (sInsUpdGroupID.ToString <> "") And (sInsUpdGroupID <> "'-12345'") Then sbw.Append("GroupID=" & sInsUpdGroupID & " and ")
      If (sInsUpdSequenceNumber.ToString <> "") And (sInsUpdSequenceNumber <> "'-12345'") Then sbw.Append("SequenceNumber=" & sInsUpdSequenceNumber & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      ID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("ID")), 0, CurrentRow.Item("ID").ToString)
      Property_ID_RQ__Integer = IIf(IsDBNull(CurrentRow.Item("Property_ID")), 0, CurrentRow.Item("Property_ID"))
      GroupID_RQ__Integer = IIf(IsDBNull(CurrentRow.Item("GroupID")), 0, CurrentRow.Item("GroupID"))
      SequenceNumber_RQ__Integer = IIf(IsDBNull(CurrentRow.Item("SequenceNumber")), 0, CurrentRow.Item("SequenceNumber"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [PropertiesGroups](")
    If sInsUpdProperty_ID.ToString <> "" Then
      sb.Append("Property_ID,")
      sbv.Append(sInsUpdProperty_ID & ",")
    End If
    If sInsUpdGroupID.ToString <> "" Then
      sb.Append("GroupID,")
      sbv.Append(sInsUpdGroupID & ",")
    End If
    If sInsUpdSequenceNumber.ToString <> "" Then
      sb.Append("SequenceNumber,")
      sbv.Append(sInsUpdSequenceNumber & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(ID) from [PropertiesGroups]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    ID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [PropertiesGroups] Set ")
    If sInsUpdProperty_ID.ToString <> "" Then sb.Append("Property_ID=" & sInsUpdProperty_ID & ",")
    If sInsUpdGroupID.ToString <> "" Then sb.Append("GroupID=" & sInsUpdGroupID & ",")
    If sInsUpdSequenceNumber.ToString <> "" Then sb.Append("SequenceNumber=" & sInsUpdSequenceNumber & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where ID=" & sInsUpdID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [PropertiesGroups] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("ID=" & sInsUpdID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TablePropertyGroups

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iGroupID As Int32
  Private sInsUpdGroupID As String
  Property GroupID_PK__Integer() As Int32
    Get
      Return iGroupID
    End Get
    Set(ByVal Value As Int32)
      iGroupID = Value
      sInsUpdGroupID = oUtil.FixParam(iGroupID, True)
    End Set
  End Property

  Private sGroupName As String
  Private sInsUpdGroupName As String
  Property GroupName_RQ__String() As String
    Get
      Return sGroupName
    End Get
    Set(ByVal Value As String)
      sGroupName = Value
      sInsUpdGroupName = oUtil.FixParam(sGroupName, False)
    End Set
  End Property

  Private sMasterText As String
  Private sInsUpdMasterText As String
  Property MasterText__String() As String
    Get
      Return sMasterText
    End Get
    Set(ByVal Value As String)
      sMasterText = Value
      sInsUpdMasterText = oUtil.FixParam(sMasterText, True)
    End Set
  End Property

  Private sDetailsText As String
  Private sInsUpdDetailsText As String
  Property DetailsText__String() As String
    Get
      Return sDetailsText
    End Get
    Set(ByVal Value As String)
      sDetailsText = Value
      sInsUpdDetailsText = oUtil.FixParam(sDetailsText, True)
    End Set
  End Property

  Public Sub Clear()
    iGroupID = 0
    sInsUpdGroupID = ""
    sGroupName = ""
    sInsUpdGroupName = ""
    sMasterText = ""
    sInsUpdMasterText = ""
    sDetailsText = ""
    sInsUpdDetailsText = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdGroupID.ToString = "'-12345'" Then sb.Append("GroupID,")
        If sInsUpdGroupName.ToString = "'Select'" Then sb.Append("GroupName,")
        If sInsUpdMasterText.ToString = "'Select'" Then sb.Append("MasterText,")
        If sInsUpdDetailsText.ToString = "'Select'" Then sb.Append("DetailsText,")
      Else
        sb.Append("GroupID,")
        sb.Append("GroupName,")
        sb.Append("MasterText,")
        sb.Append("DetailsText,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [PropertyGroups]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdGroupID.ToString <> "") And (sInsUpdGroupID <> "'-12345'") Then sbw.Append("GroupID=" & sInsUpdGroupID & " and ")
      If (sInsUpdGroupName.ToString <> "") And (sInsUpdGroupName <> "'Select'") Then sbw.Append("GroupName=" & sInsUpdGroupName & " and ")
      If (sInsUpdMasterText.ToString <> "") And (sInsUpdMasterText <> "'Select'") Then sbw.Append("MasterText=" & sInsUpdMasterText & " and ")
      If (sInsUpdDetailsText.ToString <> "") And (sInsUpdDetailsText <> "'Select'") Then sbw.Append("DetailsText=" & sInsUpdDetailsText & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      GroupID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("GroupID")), 0, CurrentRow.Item("GroupID").ToString)
      GroupName_RQ__String = IIf(IsDBNull(CurrentRow.Item("GroupName")), "", CurrentRow.Item("GroupName"))
      MasterText__String = IIf(IsDBNull(CurrentRow.Item("MasterText")), "", CurrentRow.Item("MasterText"))
      DetailsText__String = IIf(IsDBNull(CurrentRow.Item("DetailsText")), "", CurrentRow.Item("DetailsText"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [PropertyGroups](")
    If sInsUpdGroupName.ToString <> "" Then
      sb.Append("GroupName,")
      sbv.Append(sInsUpdGroupName & ",")
    End If
    If sInsUpdMasterText.ToString <> "" Then
      sb.Append("MasterText,")
      sbv.Append(sInsUpdMasterText & ",")
    End If
    If sInsUpdDetailsText.ToString <> "" Then
      sb.Append("DetailsText,")
      sbv.Append(sInsUpdDetailsText & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(GroupID) from [PropertyGroups]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    GroupID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [PropertyGroups] Set ")
    If sInsUpdGroupName.ToString <> "" Then sb.Append("GroupName=" & sInsUpdGroupName & ",")
    If sInsUpdMasterText.ToString <> "" Then sb.Append("MasterText=" & sInsUpdMasterText & ",")
    If sInsUpdDetailsText.ToString <> "" Then sb.Append("DetailsText=" & sInsUpdDetailsText & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where GroupID=" & sInsUpdGroupID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [PropertyGroups] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("GroupID=" & sInsUpdGroupID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TablePropertyImages

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iID As Int32
  Private sInsUpdID As String
  Property ID_PK__Integer() As Int32
    Get
      Return iID
    End Get
    Set(ByVal Value As Int32)
      iID = Value
      sInsUpdID = oUtil.FixParam(iID, True)
    End Set
  End Property

  Private iProperty_ID As Int32
  Private sInsUpdProperty_ID As String
  Property Property_ID__Integer() As Int32
    Get
      Return iProperty_ID
    End Get
    Set(ByVal Value As Int32)
      iProperty_ID = Value
      sInsUpdProperty_ID = oUtil.FixParam(iProperty_ID, True)
    End Set
  End Property

  Private iPropertyGroupID As Int32
  Private sInsUpdPropertyGroupID As String
  Property PropertyGroupID__Integer() As Int32
    Get
      Return iPropertyGroupID
    End Get
    Set(ByVal Value As Int32)
      iPropertyGroupID = Value
      sInsUpdPropertyGroupID = oUtil.FixParam(iPropertyGroupID, True)
    End Set
  End Property

  Private sImageFile As String
  Private sInsUpdImageFile As String
  Property ImageFile_RQ__String() As String
    Get
      Return sImageFile
    End Get
    Set(ByVal Value As String)
      sImageFile = Value
      sInsUpdImageFile = oUtil.FixParam(sImageFile, False)
    End Set
  End Property

  Private iImageSequence As Int32
  Private sInsUpdImageSequence As String
  Property ImageSequence_RQ__Integer() As Int32
    Get
      Return iImageSequence
    End Get
    Set(ByVal Value As Int32)
      iImageSequence = Value
      sInsUpdImageSequence = oUtil.FixParam(iImageSequence, False)
    End Set
  End Property

  Private sImageType As String
  Private sInsUpdImageType As String
  Property ImageType_RQ__String() As String
    Get
      Return sImageType
    End Get
    Set(ByVal Value As String)
      sImageType = Value
      sInsUpdImageType = oUtil.FixParam(sImageType, False)
    End Set
  End Property

  Private sImageStatus As String
  Private sInsUpdImageStatus As String
  Property ImageStatus_RQ__String() As String
    Get
      Return sImageStatus
    End Get
    Set(ByVal Value As String)
      sImageStatus = Value
      sInsUpdImageStatus = oUtil.FixParam(sImageStatus, False)
    End Set
  End Property

  Private sImageFileEnlarged As String
  Private sInsUpdImageFileEnlarged As String
  Property ImageFileEnlarged__String() As String
    Get
      Return sImageFileEnlarged
    End Get
    Set(ByVal Value As String)
      sImageFileEnlarged = Value
      sInsUpdImageFileEnlarged = oUtil.FixParam(sImageFileEnlarged, True)
    End Set
  End Property

  Private iImageHeight As Int32
  Private sInsUpdImageHeight As String
  Property ImageHeight__Integer() As Int32
    Get
      Return iImageHeight
    End Get
    Set(ByVal Value As Int32)
      iImageHeight = Value
      sInsUpdImageHeight = oUtil.FixParam(iImageHeight, True)
    End Set
  End Property

  Private iImageWidth As Int32
  Private sInsUpdImageWidth As String
  Property ImageWidth__Integer() As Int32
    Get
      Return iImageWidth
    End Get
    Set(ByVal Value As Int32)
      iImageWidth = Value
      sInsUpdImageWidth = oUtil.FixParam(iImageWidth, True)
    End Set
  End Property

  Private sImageAlt As String
  Private sInsUpdImageAlt As String
  Property ImageAlt__String() As String
    Get
      Return sImageAlt
    End Get
    Set(ByVal Value As String)
      sImageAlt = Value
      sInsUpdImageAlt = oUtil.FixParam(sImageAlt, True)
    End Set
  End Property

  Private sImageSectionTitle As String
  Private sInsUpdImageSectionTitle As String
  Property ImageSectionTitle__String() As String
    Get
      Return sImageSectionTitle
    End Get
    Set(ByVal Value As String)
      sImageSectionTitle = Value
      sInsUpdImageSectionTitle = oUtil.FixParam(sImageSectionTitle, True)
    End Set
  End Property

  Private sImageLocation As String
  Private sInsUpdImageLocation As String
  Property ImageLocation__String() As String
    Get
      Return sImageLocation
    End Get
    Set(ByVal Value As String)
      sImageLocation = Value
      sInsUpdImageLocation = oUtil.FixParam(sImageLocation, True)
    End Set
  End Property

  Public Sub Clear()
    iID = 0
    sInsUpdID = ""
    iProperty_ID = 0
    sInsUpdProperty_ID = ""
    iPropertyGroupID = 0
    sInsUpdPropertyGroupID = ""
    sImageFile = ""
    sInsUpdImageFile = ""
    iImageSequence = 0
    sInsUpdImageSequence = ""
    sImageType = ""
    sInsUpdImageType = ""
    sImageStatus = ""
    sInsUpdImageStatus = ""
    sImageFileEnlarged = ""
    sInsUpdImageFileEnlarged = ""
    iImageHeight = 0
    sInsUpdImageHeight = ""
    iImageWidth = 0
    sInsUpdImageWidth = ""
    sImageAlt = ""
    sInsUpdImageAlt = ""
    sImageSectionTitle = ""
    sInsUpdImageSectionTitle = ""
    sImageLocation = ""
    sInsUpdImageLocation = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdID.ToString = "'-12345'" Then sb.Append("ID,")
        If sInsUpdProperty_ID.ToString = "'-12345'" Then sb.Append("Property_ID,")
        If sInsUpdPropertyGroupID.ToString = "'-12345'" Then sb.Append("PropertyGroupID,")
        If sInsUpdImageFile.ToString = "'Select'" Then sb.Append("ImageFile,")
        If sInsUpdImageSequence.ToString = "'-12345'" Then sb.Append("ImageSequence,")
        If sInsUpdImageType.ToString = "'Select'" Then sb.Append("ImageType,")
        If sInsUpdImageStatus.ToString = "'Select'" Then sb.Append("ImageStatus,")
        If sInsUpdImageFileEnlarged.ToString = "'Select'" Then sb.Append("ImageFileEnlarged,")
        If sInsUpdImageHeight.ToString = "'-12345'" Then sb.Append("ImageHeight,")
        If sInsUpdImageWidth.ToString = "'-12345'" Then sb.Append("ImageWidth,")
        If sInsUpdImageAlt.ToString = "'Select'" Then sb.Append("ImageAlt,")
        If sInsUpdImageSectionTitle.ToString = "'Select'" Then sb.Append("ImageSectionTitle,")
        If sInsUpdImageLocation.ToString = "'Select'" Then sb.Append("ImageLocation,")
      Else
        sb.Append("ID,")
        sb.Append("Property_ID,")
        sb.Append("PropertyGroupID,")
        sb.Append("ImageFile,")
        sb.Append("ImageSequence,")
        sb.Append("ImageType,")
        sb.Append("ImageStatus,")
        sb.Append("ImageFileEnlarged,")
        sb.Append("ImageHeight,")
        sb.Append("ImageWidth,")
        sb.Append("ImageAlt,")
        sb.Append("ImageSectionTitle,")
        sb.Append("ImageLocation,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [PropertyImages]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdID.ToString <> "") And (sInsUpdID <> "'-12345'") Then sbw.Append("ID=" & sInsUpdID & " and ")
      If (sInsUpdProperty_ID.ToString <> "") And (sInsUpdProperty_ID <> "'-12345'") Then sbw.Append("Property_ID=" & sInsUpdProperty_ID & " and ")
      If (sInsUpdPropertyGroupID.ToString <> "") And (sInsUpdPropertyGroupID <> "'-12345'") Then sbw.Append("PropertyGroupID=" & sInsUpdPropertyGroupID & " and ")
      If (sInsUpdImageFile.ToString <> "") And (sInsUpdImageFile <> "'Select'") Then sbw.Append("ImageFile=" & sInsUpdImageFile & " and ")
      If (sInsUpdImageSequence.ToString <> "") And (sInsUpdImageSequence <> "'-12345'") Then sbw.Append("ImageSequence=" & sInsUpdImageSequence & " and ")
      If (sInsUpdImageType.ToString <> "") And (sInsUpdImageType <> "'Select'") Then sbw.Append("ImageType=" & sInsUpdImageType & " and ")
      If (sInsUpdImageStatus.ToString <> "") And (sInsUpdImageStatus <> "'Select'") Then sbw.Append("ImageStatus=" & sInsUpdImageStatus & " and ")
      If (sInsUpdImageFileEnlarged.ToString <> "") And (sInsUpdImageFileEnlarged <> "'Select'") Then sbw.Append("ImageFileEnlarged=" & sInsUpdImageFileEnlarged & " and ")
      If (sInsUpdImageHeight.ToString <> "") And (sInsUpdImageHeight <> "'-12345'") Then sbw.Append("ImageHeight=" & sInsUpdImageHeight & " and ")
      If (sInsUpdImageWidth.ToString <> "") And (sInsUpdImageWidth <> "'-12345'") Then sbw.Append("ImageWidth=" & sInsUpdImageWidth & " and ")
      If (sInsUpdImageAlt.ToString <> "") And (sInsUpdImageAlt <> "'Select'") Then sbw.Append("ImageAlt=" & sInsUpdImageAlt & " and ")
      If (sInsUpdImageSectionTitle.ToString <> "") And (sInsUpdImageSectionTitle <> "'Select'") Then sbw.Append("ImageSectionTitle=" & sInsUpdImageSectionTitle & " and ")
      If (sInsUpdImageLocation.ToString <> "") And (sInsUpdImageLocation <> "'Select'") Then sbw.Append("ImageLocation=" & sInsUpdImageLocation & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      ID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("ID")), 0, CurrentRow.Item("ID").ToString)
      Property_ID__Integer = IIf(IsDBNull(CurrentRow.Item("Property_ID")), 0, CurrentRow.Item("Property_ID"))
      PropertyGroupID__Integer = IIf(IsDBNull(CurrentRow.Item("PropertyGroupID")), 0, CurrentRow.Item("PropertyGroupID"))
      ImageFile_RQ__String = IIf(IsDBNull(CurrentRow.Item("ImageFile")), "", CurrentRow.Item("ImageFile"))
      ImageSequence_RQ__Integer = IIf(IsDBNull(CurrentRow.Item("ImageSequence")), 0, CurrentRow.Item("ImageSequence"))
      ImageType_RQ__String = IIf(IsDBNull(CurrentRow.Item("ImageType")), "", CurrentRow.Item("ImageType"))
      ImageStatus_RQ__String = IIf(IsDBNull(CurrentRow.Item("ImageStatus")), "", CurrentRow.Item("ImageStatus"))
      ImageFileEnlarged__String = IIf(IsDBNull(CurrentRow.Item("ImageFileEnlarged")), "", CurrentRow.Item("ImageFileEnlarged"))
      ImageHeight__Integer = IIf(IsDBNull(CurrentRow.Item("ImageHeight")), 0, CurrentRow.Item("ImageHeight"))
      ImageWidth__Integer = IIf(IsDBNull(CurrentRow.Item("ImageWidth")), 0, CurrentRow.Item("ImageWidth"))
      ImageAlt__String = IIf(IsDBNull(CurrentRow.Item("ImageAlt")), "", CurrentRow.Item("ImageAlt"))
      ImageSectionTitle__String = IIf(IsDBNull(CurrentRow.Item("ImageSectionTitle")), "", CurrentRow.Item("ImageSectionTitle"))
      ImageLocation__String = IIf(IsDBNull(CurrentRow.Item("ImageLocation")), "", CurrentRow.Item("ImageLocation"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [PropertyImages](")
    If sInsUpdProperty_ID.ToString <> "" Then
      sb.Append("Property_ID,")
      sbv.Append(sInsUpdProperty_ID & ",")
    End If
    If sInsUpdPropertyGroupID.ToString <> "" Then
      sb.Append("PropertyGroupID,")
      sbv.Append(sInsUpdPropertyGroupID & ",")
    End If
    If sInsUpdImageFile.ToString <> "" Then
      sb.Append("ImageFile,")
      sbv.Append(sInsUpdImageFile & ",")
    End If
    If sInsUpdImageSequence.ToString <> "" Then
      sb.Append("ImageSequence,")
      sbv.Append(sInsUpdImageSequence & ",")
    End If
    If sInsUpdImageType.ToString <> "" Then
      sb.Append("ImageType,")
      sbv.Append(sInsUpdImageType & ",")
    End If
    If sInsUpdImageStatus.ToString <> "" Then
      sb.Append("ImageStatus,")
      sbv.Append(sInsUpdImageStatus & ",")
    End If
    If sInsUpdImageFileEnlarged.ToString <> "" Then
      sb.Append("ImageFileEnlarged,")
      sbv.Append(sInsUpdImageFileEnlarged & ",")
    End If
    If sInsUpdImageHeight.ToString <> "" Then
      sb.Append("ImageHeight,")
      sbv.Append(sInsUpdImageHeight & ",")
    End If
    If sInsUpdImageWidth.ToString <> "" Then
      sb.Append("ImageWidth,")
      sbv.Append(sInsUpdImageWidth & ",")
    End If
    If sInsUpdImageAlt.ToString <> "" Then
      sb.Append("ImageAlt,")
      sbv.Append(sInsUpdImageAlt & ",")
    End If
    If sInsUpdImageSectionTitle.ToString <> "" Then
      sb.Append("ImageSectionTitle,")
      sbv.Append(sInsUpdImageSectionTitle & ",")
    End If
    If sInsUpdImageLocation.ToString <> "" Then
      sb.Append("ImageLocation,")
      sbv.Append(sInsUpdImageLocation & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(ID) from [PropertyImages]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    ID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [PropertyImages] Set ")
    If sInsUpdProperty_ID.ToString <> "" Then sb.Append("Property_ID=" & sInsUpdProperty_ID & ",")
    If sInsUpdPropertyGroupID.ToString <> "" Then sb.Append("PropertyGroupID=" & sInsUpdPropertyGroupID & ",")
    If sInsUpdImageFile.ToString <> "" Then sb.Append("ImageFile=" & sInsUpdImageFile & ",")
    If sInsUpdImageSequence.ToString <> "" Then sb.Append("ImageSequence=" & sInsUpdImageSequence & ",")
    If sInsUpdImageType.ToString <> "" Then sb.Append("ImageType=" & sInsUpdImageType & ",")
    If sInsUpdImageStatus.ToString <> "" Then sb.Append("ImageStatus=" & sInsUpdImageStatus & ",")
    If sInsUpdImageFileEnlarged.ToString <> "" Then sb.Append("ImageFileEnlarged=" & sInsUpdImageFileEnlarged & ",")
    If sInsUpdImageHeight.ToString <> "" Then sb.Append("ImageHeight=" & sInsUpdImageHeight & ",")
    If sInsUpdImageWidth.ToString <> "" Then sb.Append("ImageWidth=" & sInsUpdImageWidth & ",")
    If sInsUpdImageAlt.ToString <> "" Then sb.Append("ImageAlt=" & sInsUpdImageAlt & ",")
    If sInsUpdImageSectionTitle.ToString <> "" Then sb.Append("ImageSectionTitle=" & sInsUpdImageSectionTitle & ",")
    If sInsUpdImageLocation.ToString <> "" Then sb.Append("ImageLocation=" & sInsUpdImageLocation & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where ID=" & sInsUpdID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [PropertyImages] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("ID=" & sInsUpdID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TableQueuedEmail

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iID As Int32
  Private sInsUpdID As String
  Property ID_RQ__Integer() As Int32
    Get
      Return iID
    End Get
    Set(ByVal Value As Int32)
      iID = Value
      sInsUpdID = oUtil.FixParam(iID, False)
    End Set
  End Property

  Private sInsUpdOrigID As String
  Private iOrigID As Int32
  Property ID_RQ__Integer_Orig() As Int32
    Get
      Return iOrigID
    End Get
    Set(ByVal Value As Int32)
      iOrigID = Value
      sInsUpdOrigID = oUtil.FixParam(iOrigID, False)
    End Set
  End Property

  Private sSourceServer As String
  Private sInsUpdSourceServer As String
  Property SourceServer__String() As String
    Get
      Return sSourceServer
    End Get
    Set(ByVal Value As String)
      sSourceServer = Value
      sInsUpdSourceServer = oUtil.FixParam(sSourceServer, True)
    End Set
  End Property

  Private sSourceDB As String
  Private sInsUpdSourceDB As String
  Property SourceDB__String() As String
    Get
      Return sSourceDB
    End Get
    Set(ByVal Value As String)
      sSourceDB = Value
      sInsUpdSourceDB = oUtil.FixParam(sSourceDB, True)
    End Set
  End Property

  Private sLinkedTable As String
  Private sInsUpdLinkedTable As String
  Property LinkedTable__String() As String
    Get
      Return sLinkedTable
    End Get
    Set(ByVal Value As String)
      sLinkedTable = Value
      sInsUpdLinkedTable = oUtil.FixParam(sLinkedTable, True)
    End Set
  End Property

  Private iLinkedTableID As Int32
  Private sInsUpdLinkedTableID As String
  Property LinkedTableID__Integer() As Int32
    Get
      Return iLinkedTableID
    End Get
    Set(ByVal Value As Int32)
      iLinkedTableID = Value
      sInsUpdLinkedTableID = oUtil.FixParam(iLinkedTableID, True)
    End Set
  End Property

  Private iFormatIsHTML As Int32
  Private sInsUpdFormatIsHTML As String
  Property FormatIsHTML_RQ__Integer() As Int32
    Get
      Return iFormatIsHTML
    End Get
    Set(ByVal Value As Int32)
      iFormatIsHTML = Value
      sInsUpdFormatIsHTML = oUtil.FixParam(iFormatIsHTML, False)
    End Set
  End Property

  Private sInsUpdOrigFormatIsHTML As String
  Private iOrigFormatIsHTML As Int32
  Property FormatIsHTML_RQ__Integer_Orig() As Int32
    Get
      Return iOrigFormatIsHTML
    End Get
    Set(ByVal Value As Int32)
      iOrigFormatIsHTML = Value
      sInsUpdOrigFormatIsHTML = oUtil.FixParam(iOrigFormatIsHTML, False)
    End Set
  End Property

  Private sFromAddress As String
  Private sInsUpdFromAddress As String
  Property FromAddress_RQ__String() As String
    Get
      Return sFromAddress
    End Get
    Set(ByVal Value As String)
      sFromAddress = Value
      sInsUpdFromAddress = oUtil.FixParam(sFromAddress, False)
    End Set
  End Property

  Private sInsUpdOrigFromAddress As String
  Private sOrigFromAddress As String
  Property FromAddress_RQ__String_Orig() As String
    Get
      Return sOrigFromAddress
    End Get
    Set(ByVal Value As String)
      sOrigFromAddress = Value
      sInsUpdOrigFromAddress = oUtil.FixParam(sOrigFromAddress, False)
    End Set
  End Property

  Private sFromName As String
  Private sInsUpdFromName As String
  Property FromName__String() As String
    Get
      Return sFromName
    End Get
    Set(ByVal Value As String)
      sFromName = Value
      sInsUpdFromName = oUtil.FixParam(sFromName, True)
    End Set
  End Property

  Private sToAddress As String
  Private sInsUpdToAddress As String
  Property ToAddress_RQ__String() As String
    Get
      Return sToAddress
    End Get
    Set(ByVal Value As String)
      sToAddress = Value
      sInsUpdToAddress = oUtil.FixParam(sToAddress, False)
    End Set
  End Property

  Private sInsUpdOrigToAddress As String
  Private sOrigToAddress As String
  Property ToAddress_RQ__String_Orig() As String
    Get
      Return sOrigToAddress
    End Get
    Set(ByVal Value As String)
      sOrigToAddress = Value
      sInsUpdOrigToAddress = oUtil.FixParam(sOrigToAddress, False)
    End Set
  End Property

  Private sToName As String
  Private sInsUpdToName As String
  Property ToName__String() As String
    Get
      Return sToName
    End Get
    Set(ByVal Value As String)
      sToName = Value
      sInsUpdToName = oUtil.FixParam(sToName, True)
    End Set
  End Property

  Private sSubject As String
  Private sInsUpdSubject As String
  Property Subject_RQ__String() As String
    Get
      Return sSubject
    End Get
    Set(ByVal Value As String)
      sSubject = Value
      sInsUpdSubject = oUtil.FixParam(sSubject, False)
    End Set
  End Property

  Private sInsUpdOrigSubject As String
  Private sOrigSubject As String
  Property Subject_RQ__String_Orig() As String
    Get
      Return sOrigSubject
    End Get
    Set(ByVal Value As String)
      sOrigSubject = Value
      sInsUpdOrigSubject = oUtil.FixParam(sOrigSubject, False)
    End Set
  End Property

  Private sBody As String
  Private sInsUpdBody As String
  Property Body_RQ__String() As String
    Get
      Return sBody
    End Get
    Set(ByVal Value As String)
      sBody = Value
      sInsUpdBody = oUtil.FixParam(sBody, False)
    End Set
  End Property

  Private sInsUpdOrigBody As String
  Private sOrigBody As String
  Property Body_RQ__String_Orig() As String
    Get
      Return sOrigBody
    End Get
    Set(ByVal Value As String)
      sOrigBody = Value
      sInsUpdOrigBody = oUtil.FixParam(sOrigBody, False)
    End Set
  End Property

  Private sAttachmentFile As String
  Private sInsUpdAttachmentFile As String
  Property AttachmentFile__String() As String
    Get
      Return sAttachmentFile
    End Get
    Set(ByVal Value As String)
      sAttachmentFile = Value
      sInsUpdAttachmentFile = oUtil.FixParam(sAttachmentFile, True)
    End Set
  End Property

  Private sStatus As String
  Private sInsUpdStatus As String
  Property Status_RQ__String() As String
    Get
      Return sStatus
    End Get
    Set(ByVal Value As String)
      sStatus = Value
      sInsUpdStatus = oUtil.FixParam(sStatus, False)
    End Set
  End Property

  Private sInsUpdOrigStatus As String
  Private sOrigStatus As String
  Property Status_RQ__String_Orig() As String
    Get
      Return sOrigStatus
    End Get
    Set(ByVal Value As String)
      sOrigStatus = Value
      sInsUpdOrigStatus = oUtil.FixParam(sOrigStatus, False)
    End Set
  End Property

  Private sAddDate As String
  Private sInsUpdAddDate As String
  Property AddDate__Date() As String
    Get
      Return sAddDate
    End Get
    Set(ByVal Value As String)
      sAddDate = Value
      sInsUpdAddDate = oUtil.FixParam(sAddDate, True)
    End Set
  End Property

  Private sSendDate As String
  Private sInsUpdSendDate As String
  Property SendDate__Date() As String
    Get
      Return sSendDate
    End Get
    Set(ByVal Value As String)
      sSendDate = Value
      sInsUpdSendDate = oUtil.FixParam(sSendDate, True)
    End Set
  End Property

  Private iPriority As Int32
  Private sInsUpdPriority As String
  Property Priority__Integer() As Int32
    Get
      Return iPriority
    End Get
    Set(ByVal Value As Int32)
      iPriority = Value
      sInsUpdPriority = oUtil.FixParam(iPriority, True)
    End Set
  End Property

  Private iTries As Int32
  Private sInsUpdTries As String
  Property Tries__Integer() As Int32
    Get
      Return iTries
    End Get
    Set(ByVal Value As Int32)
      iTries = Value
      sInsUpdTries = oUtil.FixParam(iTries, True)
    End Set
  End Property

  Private sErrMessage As String
  Private sInsUpdErrMessage As String
  Property ErrMessage__String() As String
    Get
      Return sErrMessage
    End Get
    Set(ByVal Value As String)
      sErrMessage = Value
      sInsUpdErrMessage = oUtil.FixParam(sErrMessage, True)
    End Set
  End Property

  Private sSubmittedBy As String
  Private sInsUpdSubmittedBy As String
  Property SubmittedBy__String() As String
    Get
      Return sSubmittedBy
    End Get
    Set(ByVal Value As String)
      sSubmittedBy = Value
      sInsUpdSubmittedBy = oUtil.FixParam(sSubmittedBy, True)
    End Set
  End Property

  Public Sub Clear()
    iID = 0
    sInsUpdID = ""
    iOrigID = 0
    sInsUpdOrigID = ""
    sSourceServer = ""
    sInsUpdSourceServer = ""
    sSourceDB = ""
    sInsUpdSourceDB = ""
    sLinkedTable = ""
    sInsUpdLinkedTable = ""
    iLinkedTableID = 0
    sInsUpdLinkedTableID = ""
    iFormatIsHTML = 0
    sInsUpdFormatIsHTML = ""
    iOrigFormatIsHTML = 0
    sInsUpdOrigFormatIsHTML = ""
    sFromAddress = ""
    sInsUpdFromAddress = ""
    sOrigFromAddress = ""
    sInsUpdOrigFromAddress = ""
    sFromName = ""
    sInsUpdFromName = ""
    sToAddress = ""
    sInsUpdToAddress = ""
    sOrigToAddress = ""
    sInsUpdOrigToAddress = ""
    sToName = ""
    sInsUpdToName = ""
    sSubject = ""
    sInsUpdSubject = ""
    sOrigSubject = ""
    sInsUpdOrigSubject = ""
    sBody = ""
    sInsUpdBody = ""
    sOrigBody = ""
    sInsUpdOrigBody = ""
    sAttachmentFile = ""
    sInsUpdAttachmentFile = ""
    sStatus = ""
    sInsUpdStatus = ""
    sOrigStatus = ""
    sInsUpdOrigStatus = ""
    sAddDate = ""
    sInsUpdAddDate = ""
    sSendDate = ""
    sInsUpdSendDate = ""
    iPriority = 0
    sInsUpdPriority = ""
    iTries = 0
    sInsUpdTries = ""
    sErrMessage = ""
    sInsUpdErrMessage = ""
    sSubmittedBy = ""
    sInsUpdSubmittedBy = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdID.ToString = "'-12345'" Then sb.Append("ID,")
        If sInsUpdSourceServer.ToString = "'Select'" Then sb.Append("SourceServer,")
        If sInsUpdSourceDB.ToString = "'Select'" Then sb.Append("SourceDB,")
        If sInsUpdLinkedTable.ToString = "'Select'" Then sb.Append("LinkedTable,")
        If sInsUpdLinkedTableID.ToString = "'-12345'" Then sb.Append("LinkedTableID,")
        If sInsUpdFormatIsHTML.ToString = "'-12345'" Then sb.Append("FormatIsHTML,")
        If sInsUpdFromAddress.ToString = "'Select'" Then sb.Append("FromAddress,")
        If sInsUpdFromName.ToString = "'Select'" Then sb.Append("FromName,")
        If sInsUpdToAddress.ToString = "'Select'" Then sb.Append("ToAddress,")
        If sInsUpdToName.ToString = "'Select'" Then sb.Append("ToName,")
        If sInsUpdSubject.ToString = "'Select'" Then sb.Append("Subject,")
        If sInsUpdBody.ToString = "'Select'" Then sb.Append("Body,")
        If sInsUpdAttachmentFile.ToString = "'Select'" Then sb.Append("AttachmentFile,")
        If sInsUpdStatus.ToString = "'Select'" Then sb.Append("Status,")
        If sInsUpdAddDate.ToString = "'Select'" Then sb.Append("AddDate,")
        If sInsUpdSendDate.ToString = "'Select'" Then sb.Append("SendDate,")
        If sInsUpdPriority.ToString = "'-12345'" Then sb.Append("Priority,")
        If sInsUpdTries.ToString = "'-12345'" Then sb.Append("Tries,")
        If sInsUpdErrMessage.ToString = "'Select'" Then sb.Append("ErrMessage,")
        If sInsUpdSubmittedBy.ToString = "'Select'" Then sb.Append("SubmittedBy,")
      Else
        sb.Append("ID,")
        sb.Append("SourceServer,")
        sb.Append("SourceDB,")
        sb.Append("LinkedTable,")
        sb.Append("LinkedTableID,")
        sb.Append("FormatIsHTML,")
        sb.Append("FromAddress,")
        sb.Append("FromName,")
        sb.Append("ToAddress,")
        sb.Append("ToName,")
        sb.Append("Subject,")
        sb.Append("Body,")
        sb.Append("AttachmentFile,")
        sb.Append("Status,")
        sb.Append("AddDate,")
        sb.Append("SendDate,")
        sb.Append("Priority,")
        sb.Append("Tries,")
        sb.Append("ErrMessage,")
        sb.Append("SubmittedBy,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [QueuedEmail]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdID.ToString <> "") And (sInsUpdID <> "'-12345'") Then sbw.Append("ID=" & sInsUpdID & " and ")
      If (sInsUpdSourceServer.ToString <> "") And (sInsUpdSourceServer <> "'Select'") Then sbw.Append("SourceServer=" & sInsUpdSourceServer & " and ")
      If (sInsUpdSourceDB.ToString <> "") And (sInsUpdSourceDB <> "'Select'") Then sbw.Append("SourceDB=" & sInsUpdSourceDB & " and ")
      If (sInsUpdLinkedTable.ToString <> "") And (sInsUpdLinkedTable <> "'Select'") Then sbw.Append("LinkedTable=" & sInsUpdLinkedTable & " and ")
      If (sInsUpdLinkedTableID.ToString <> "") And (sInsUpdLinkedTableID <> "'-12345'") Then sbw.Append("LinkedTableID=" & sInsUpdLinkedTableID & " and ")
      If (sInsUpdFormatIsHTML.ToString <> "") And (sInsUpdFormatIsHTML <> "'-12345'") Then sbw.Append("FormatIsHTML=" & sInsUpdFormatIsHTML & " and ")
      If (sInsUpdFromAddress.ToString <> "") And (sInsUpdFromAddress <> "'Select'") Then sbw.Append("FromAddress=" & sInsUpdFromAddress & " and ")
      If (sInsUpdFromName.ToString <> "") And (sInsUpdFromName <> "'Select'") Then sbw.Append("FromName=" & sInsUpdFromName & " and ")
      If (sInsUpdToAddress.ToString <> "") And (sInsUpdToAddress <> "'Select'") Then sbw.Append("ToAddress=" & sInsUpdToAddress & " and ")
      If (sInsUpdToName.ToString <> "") And (sInsUpdToName <> "'Select'") Then sbw.Append("ToName=" & sInsUpdToName & " and ")
      If (sInsUpdSubject.ToString <> "") And (sInsUpdSubject <> "'Select'") Then sbw.Append("Subject=" & sInsUpdSubject & " and ")
      If (sInsUpdBody.ToString <> "") And (sInsUpdBody <> "'Select'") Then sbw.Append("Body=" & sInsUpdBody & " and ")
      If (sInsUpdAttachmentFile.ToString <> "") And (sInsUpdAttachmentFile <> "'Select'") Then sbw.Append("AttachmentFile=" & sInsUpdAttachmentFile & " and ")
      If (sInsUpdStatus.ToString <> "") And (sInsUpdStatus <> "'Select'") Then sbw.Append("Status=" & sInsUpdStatus & " and ")
      If (sInsUpdAddDate.ToString <> "") And (sInsUpdAddDate <> "'Select'") Then sbw.Append("AddDate=" & sInsUpdAddDate & " and ")
      If (sInsUpdSendDate.ToString <> "") And (sInsUpdSendDate <> "'Select'") Then sbw.Append("SendDate=" & sInsUpdSendDate & " and ")
      If (sInsUpdPriority.ToString <> "") And (sInsUpdPriority <> "'-12345'") Then sbw.Append("Priority=" & sInsUpdPriority & " and ")
      If (sInsUpdTries.ToString <> "") And (sInsUpdTries <> "'-12345'") Then sbw.Append("Tries=" & sInsUpdTries & " and ")
      If (sInsUpdErrMessage.ToString <> "") And (sInsUpdErrMessage <> "'Select'") Then sbw.Append("ErrMessage=" & sInsUpdErrMessage & " and ")
      If (sInsUpdSubmittedBy.ToString <> "") And (sInsUpdSubmittedBy <> "'Select'") Then sbw.Append("SubmittedBy=" & sInsUpdSubmittedBy & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      ID_RQ__Integer = IIf(IsDBNull(CurrentRow.Item("ID")), 0, CurrentRow.Item("ID"))
      SourceServer__String = IIf(IsDBNull(CurrentRow.Item("SourceServer")), "", CurrentRow.Item("SourceServer"))
      SourceDB__String = IIf(IsDBNull(CurrentRow.Item("SourceDB")), "", CurrentRow.Item("SourceDB"))
      LinkedTable__String = IIf(IsDBNull(CurrentRow.Item("LinkedTable")), "", CurrentRow.Item("LinkedTable"))
      LinkedTableID__Integer = IIf(IsDBNull(CurrentRow.Item("LinkedTableID")), 0, CurrentRow.Item("LinkedTableID"))
      FormatIsHTML_RQ__Integer = IIf(IsDBNull(CurrentRow.Item("FormatIsHTML")), 0, CurrentRow.Item("FormatIsHTML"))
      FromAddress_RQ__String = IIf(IsDBNull(CurrentRow.Item("FromAddress")), "", CurrentRow.Item("FromAddress"))
      FromName__String = IIf(IsDBNull(CurrentRow.Item("FromName")), "", CurrentRow.Item("FromName"))
      ToAddress_RQ__String = IIf(IsDBNull(CurrentRow.Item("ToAddress")), "", CurrentRow.Item("ToAddress"))
      ToName__String = IIf(IsDBNull(CurrentRow.Item("ToName")), "", CurrentRow.Item("ToName"))
      Subject_RQ__String = IIf(IsDBNull(CurrentRow.Item("Subject")), "", CurrentRow.Item("Subject"))
      Body_RQ__String = IIf(IsDBNull(CurrentRow.Item("Body")), "", CurrentRow.Item("Body"))
      AttachmentFile__String = IIf(IsDBNull(CurrentRow.Item("AttachmentFile")), "", CurrentRow.Item("AttachmentFile"))
      Status_RQ__String = IIf(IsDBNull(CurrentRow.Item("Status")), "", CurrentRow.Item("Status"))
      AddDate__Date = IIf(IsDBNull(CurrentRow.Item("AddDate")), "", CurrentRow.Item("AddDate"))
      SendDate__Date = IIf(IsDBNull(CurrentRow.Item("SendDate")), "", CurrentRow.Item("SendDate"))
      Priority__Integer = IIf(IsDBNull(CurrentRow.Item("Priority")), 0, CurrentRow.Item("Priority"))
      Tries__Integer = IIf(IsDBNull(CurrentRow.Item("Tries")), 0, CurrentRow.Item("Tries"))
      ErrMessage__String = IIf(IsDBNull(CurrentRow.Item("ErrMessage")), "", CurrentRow.Item("ErrMessage"))
      SubmittedBy__String = IIf(IsDBNull(CurrentRow.Item("SubmittedBy")), "", CurrentRow.Item("SubmittedBy"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [QueuedEmail](")
    If sInsUpdSourceServer.ToString <> "" Then
      sb.Append("SourceServer,")
      sbv.Append(sInsUpdSourceServer & ",")
    End If
    If sInsUpdSourceDB.ToString <> "" Then
      sb.Append("SourceDB,")
      sbv.Append(sInsUpdSourceDB & ",")
    End If
    If sInsUpdLinkedTable.ToString <> "" Then
      sb.Append("LinkedTable,")
      sbv.Append(sInsUpdLinkedTable & ",")
    End If
    If sInsUpdLinkedTableID.ToString <> "" Then
      sb.Append("LinkedTableID,")
      sbv.Append(sInsUpdLinkedTableID & ",")
    End If
    If sInsUpdFormatIsHTML.ToString <> "" Then
      sb.Append("FormatIsHTML,")
      sbv.Append(sInsUpdFormatIsHTML & ",")
    End If
    If sInsUpdFromAddress.ToString <> "" Then
      sb.Append("FromAddress,")
      sbv.Append(sInsUpdFromAddress & ",")
    End If
    If sInsUpdFromName.ToString <> "" Then
      sb.Append("FromName,")
      sbv.Append(sInsUpdFromName & ",")
    End If
    If sInsUpdToAddress.ToString <> "" Then
      sb.Append("ToAddress,")
      sbv.Append(sInsUpdToAddress & ",")
    End If
    If sInsUpdToName.ToString <> "" Then
      sb.Append("ToName,")
      sbv.Append(sInsUpdToName & ",")
    End If
    If sInsUpdSubject.ToString <> "" Then
      sb.Append("Subject,")
      sbv.Append(sInsUpdSubject & ",")
    End If
    If sInsUpdBody.ToString <> "" Then
      sb.Append("Body,")
      sbv.Append(sInsUpdBody & ",")
    End If
    If sInsUpdAttachmentFile.ToString <> "" Then
      sb.Append("AttachmentFile,")
      sbv.Append(sInsUpdAttachmentFile & ",")
    End If
    If sInsUpdStatus.ToString <> "" Then
      sb.Append("Status,")
      sbv.Append(sInsUpdStatus & ",")
    End If
    If sInsUpdAddDate.ToString <> "" Then
      sb.Append("AddDate,")
      sbv.Append(sInsUpdAddDate & ",")
    End If
    If sInsUpdSendDate.ToString <> "" Then
      sb.Append("SendDate,")
      sbv.Append(sInsUpdSendDate & ",")
    End If
    If sInsUpdPriority.ToString <> "" Then
      sb.Append("Priority,")
      sbv.Append(sInsUpdPriority & ",")
    End If
    If sInsUpdTries.ToString <> "" Then
      sb.Append("Tries,")
      sbv.Append(sInsUpdTries & ",")
    End If
    If sInsUpdErrMessage.ToString <> "" Then
      sb.Append("ErrMessage,")
      sbv.Append(sInsUpdErrMessage & ",")
    End If
    If sInsUpdSubmittedBy.ToString <> "" Then
      sb.Append("SubmittedBy,")
      sbv.Append(sInsUpdSubmittedBy & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(ID) from [QueuedEmail]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [QueuedEmail] Set ")
    If sInsUpdSourceServer.ToString <> "" Then sb.Append("SourceServer=" & sInsUpdSourceServer & ",")
    If sInsUpdSourceDB.ToString <> "" Then sb.Append("SourceDB=" & sInsUpdSourceDB & ",")
    If sInsUpdLinkedTable.ToString <> "" Then sb.Append("LinkedTable=" & sInsUpdLinkedTable & ",")
    If sInsUpdLinkedTableID.ToString <> "" Then sb.Append("LinkedTableID=" & sInsUpdLinkedTableID & ",")
    If sInsUpdFormatIsHTML.ToString <> "" Then sb.Append("FormatIsHTML=" & sInsUpdFormatIsHTML & ",")
    If sInsUpdFromAddress.ToString <> "" Then sb.Append("FromAddress=" & sInsUpdFromAddress & ",")
    If sInsUpdFromName.ToString <> "" Then sb.Append("FromName=" & sInsUpdFromName & ",")
    If sInsUpdToAddress.ToString <> "" Then sb.Append("ToAddress=" & sInsUpdToAddress & ",")
    If sInsUpdToName.ToString <> "" Then sb.Append("ToName=" & sInsUpdToName & ",")
    If sInsUpdSubject.ToString <> "" Then sb.Append("Subject=" & sInsUpdSubject & ",")
    If sInsUpdBody.ToString <> "" Then sb.Append("Body=" & sInsUpdBody & ",")
    If sInsUpdAttachmentFile.ToString <> "" Then sb.Append("AttachmentFile=" & sInsUpdAttachmentFile & ",")
    If sInsUpdStatus.ToString <> "" Then sb.Append("Status=" & sInsUpdStatus & ",")
    If sInsUpdAddDate.ToString <> "" Then sb.Append("AddDate=" & sInsUpdAddDate & ",")
    If sInsUpdSendDate.ToString <> "" Then sb.Append("SendDate=" & sInsUpdSendDate & ",")
    If sInsUpdPriority.ToString <> "" Then sb.Append("Priority=" & sInsUpdPriority & ",")
    If sInsUpdTries.ToString <> "" Then sb.Append("Tries=" & sInsUpdTries & ",")
    If sInsUpdErrMessage.ToString <> "" Then sb.Append("ErrMessage=" & sInsUpdErrMessage & ",")
    If sInsUpdSubmittedBy.ToString <> "" Then sb.Append("SubmittedBy=" & sInsUpdSubmittedBy & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where "
      sSQL = sSQL & " ID=" & sInsUpdOrigID & " and "
      sSQL = sSQL & " FormatIsHTML=" & sInsUpdOrigFormatIsHTML & " and "
      sSQL = sSQL & " FromAddress=" & sInsUpdOrigFromAddress & " and "
      sSQL = sSQL & " ToAddress=" & sInsUpdOrigToAddress & " and "
      sSQL = sSQL & " Subject=" & sInsUpdOrigSubject & " and "
      sSQL = sSQL & " Body=" & sInsUpdOrigBody & " and "
      sSQL = sSQL & " Status=" & sInsUpdOrigStatus & " and "
      sSQL = Left(sSQL, Len(sSQL) - 5)

    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [QueuedEmail] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sSQL = sb.ToString
      sSQL = sSQL & " Where "
      sSQL = sSQL & " ID=" & sInsUpdOrigID & " and "
      sSQL = sSQL & " FormatIsHTML=" & sInsUpdOrigFormatIsHTML & " and "
      sSQL = sSQL & " FromAddress=" & sInsUpdOrigFromAddress & " and "
      sSQL = sSQL & " ToAddress=" & sInsUpdOrigToAddress & " and "
      sSQL = sSQL & " Subject=" & sInsUpdOrigSubject & " and "
      sSQL = sSQL & " Body=" & sInsUpdOrigBody & " and "
      sSQL = sSQL & " Status=" & sInsUpdOrigStatus & " and "
      sSQL = Left(sSQL, Len(sSQL) - 5)

    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class

Public Class TableCodeData

  Public Connection As New System.Data.SqlClient.SqlConnection()
  Public Transaction As System.Data.SqlClient.SqlTransaction
  Public SelectedData As Object
  Public CurrentRow As Object
  Public ConnectionString As String = ""
  Public CurrentRecordNumber As Integer = 0
  Public oUtil As DBUtilities
  Public Sub New(Optional ByVal bBeginTransaction As Boolean = False)

    oUtil = New DBUtilities
    ConnectionString = oUtil.CreateConnectionStringFromConfig()
    If ConnectionString.ToString = "" Then
      ConnectionString = oUtil.CNullS(System.Configuration.ConfigurationSettings.AppSettings("ConnectionString"))
    End If
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByVal sConnnectionString As String, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    ConnectionString = sConnnectionString
    Connection.ConnectionString = ConnectionString
    If bBeginTransaction Then
      oUtil.OpenConnection(Connection, Transaction, ConnectionString)
      Transaction = Connection.BeginTransaction
    End If
    Clear()
  End Sub

  Public Sub New(ByRef DBSQLConnection As System.Data.SqlClient.SqlConnection, Optional ByVal bBeginTransaction As Boolean = False)
    oUtil = New DBUtilities
    Connection = DBSQLConnection
    Clear()

    ConnectionString = DBSQLConnection.ConnectionString
    Clear()

    If bBeginTransaction Then
      Transaction = Connection.BeginTransaction
    End If
  End Sub

  Public Sub New(ByRef DBTransaction As System.Data.SqlClient.SqlTransaction)
    oUtil = New DBUtilities
    Connection = DBTransaction.Connection
    Clear()

    Transaction = DBTransaction
  End Sub

  Private iCodeID As Int32
  Private sInsUpdCodeID As String
  Property CodeID_PK__Integer() As Int32
    Get
      Return iCodeID
    End Get
    Set(ByVal Value As Int32)
      iCodeID = Value
      sInsUpdCodeID = oUtil.FixParam(iCodeID, True)
    End Set
  End Property

  Private sCodeKey As String
  Private sInsUpdCodeKey As String
  Property CodeKey__String() As String
    Get
      Return sCodeKey
    End Get
    Set(ByVal Value As String)
      sCodeKey = Value
      sInsUpdCodeKey = oUtil.FixParam(sCodeKey, True)
    End Set
  End Property

  Private sCodeValue As String
  Private sInsUpdCodeValue As String
  Property CodeValue__String() As String
    Get
      Return sCodeValue
    End Get
    Set(ByVal Value As String)
      sCodeValue = Value
      sInsUpdCodeValue = oUtil.FixParam(sCodeValue, True)
    End Set
  End Property

  Private sCodeValueLarge As String
  Private sInsUpdCodeValueLarge As String
  Property CodeValueLarge__String() As String
    Get
      Return sCodeValueLarge
    End Get
    Set(ByVal Value As String)
      sCodeValueLarge = Value
      sInsUpdCodeValueLarge = oUtil.FixParam(sCodeValueLarge, True)
    End Set
  End Property

  Public Sub Clear()
    iCodeID = 0
    sInsUpdCodeID = ""
    sCodeKey = ""
    sInsUpdCodeKey = ""
    sCodeValue = ""
    sInsUpdCodeValue = ""
    sCodeValueLarge = ""
    sInsUpdCodeValueLarge = ""
  End Sub

  Public Function SelectData(
Optional ByVal bReturnDataInProperties As Boolean = True,
Optional ByVal bReturnOnlyFirstRecord As Boolean = True,
Optional ByRef bUseDataView As Boolean = True,
Optional ByVal bUseFieldInWhereClauseIfPropertyValueSet As Boolean = True,
Optional ByVal bSelectFieldIfPropertyValueSetToSelect As Boolean = False,
Optional ByVal sSelectClause As String = "",
Optional ByVal sWhereClause As String = "",
Optional ByVal sOrderByClause As String = "") As Object

    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbw As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable()
    SelectData = 0
    If sSelectClause.ToString = "" Then
      sb.Append("Select ")
      If bSelectFieldIfPropertyValueSetToSelect And (Not bReturnDataInProperties) Then
        If sInsUpdCodeID.ToString = "'-12345'" Then sb.Append("CodeID,")
        If sInsUpdCodeKey.ToString = "'Select'" Then sb.Append("CodeKey,")
        If sInsUpdCodeValue.ToString = "'Select'" Then sb.Append("CodeValue,")
        If sInsUpdCodeValueLarge.ToString = "'Select'" Then sb.Append("CodeValueLarge,")
      Else
        sb.Append("CodeID,")
        sb.Append("CodeKey,")
        sb.Append("CodeValue,")
        sb.Append("CodeValueLarge,")
      End If

      sSQL = sb.ToString
      If Right(sSQL, 1) = "," Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
      End If

    Else
      sSQL = "Select " & sSelectClause.ToString
    End If

    sSQL = sSQL & " from [CodeData]"

    If bUseFieldInWhereClauseIfPropertyValueSet Then
      If (sInsUpdCodeID.ToString <> "") And (sInsUpdCodeID <> "'-12345'") Then sbw.Append("CodeID=" & sInsUpdCodeID & " and ")
      If (sInsUpdCodeKey.ToString <> "") And (sInsUpdCodeKey <> "'Select'") Then sbw.Append("CodeKey=" & sInsUpdCodeKey & " and ")
      If (sInsUpdCodeValue.ToString <> "") And (sInsUpdCodeValue <> "'Select'") Then sbw.Append("CodeValue=" & sInsUpdCodeValue & " and ")
      If (sInsUpdCodeValueLarge.ToString <> "") And (sInsUpdCodeValueLarge <> "'Select'") Then sbw.Append("CodeValueLarge=" & sInsUpdCodeValueLarge & " and ")
    End If

    If sWhereClause.ToString <> "" Then
      sbw.Append(sWhereClause.ToString & " and ")
    End If

    If sbw.ToString <> "" Then
      sSQL = sSQL & " Where " & Left(sbw.ToString, Len(sbw.ToString) - 4)
    End If

    If sOrderByClause.ToString <> "" Then
      sSQL = sSQL & " Order By " & sOrderByClause.ToString
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd = New System.Data.SqlClient.SqlCommand(sSQL, Connection)
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.close()
      End If
    End If

    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else

      SelectedData = oCmd.ExecuteReader
    End If

    If bReturnDataInProperties Then
      CurrentRecordNumber = -1
      Move(bReturnOnlyFirstRecord)
    Else
      SelectData = SelectedData
    End If

    If bUseDataView Then oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbw = Nothing
    oCmd = Nothing
    oSQLAdapter = Nothing
    oDataTable = Nothing
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    Clear()

    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          If sFilterForDataView.ToUpper = "NONE" Then sFilterForDataView = ""
          SelectedData.RowFilter = sFilterForDataView.ToString
          If SelectedData.Count = 0 Then Exit Function
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      CodeID_PK__Integer = IIf(IsDBNull(CurrentRow.Item("CodeID")), 0, CurrentRow.Item("CodeID").ToString)
      CodeKey__String = IIf(IsDBNull(CurrentRow.Item("CodeKey")), "", CurrentRow.Item("CodeKey"))
      CodeValue__String = IIf(IsDBNull(CurrentRow.Item("CodeValue")), "", CurrentRow.Item("CodeValue"))
      CodeValueLarge__String = IIf(IsDBNull(CurrentRow.Item("CodeValueLarge")), "", CurrentRow.Item("CodeValueLarge"))

      Move = True
      If bCloseDataSourceAfterRead And Transaction Is Nothing Then oUtil.CloseConnection(Connection, Transaction)
    End If

  End Function
  Public Sub OpenConnection()
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
  End Sub
  Public Sub CloseConnection()
    oUtil.CloseConnection(Connection, Transaction)
  End Sub
  Public Sub ProcessTransaction(Optional ByVal bCommit As Boolean = True)
    oUtil.ProcessTransaction(Connection, Transaction, bCommit)
  End Sub
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, sFilterForDataView, -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, sFilterForDataView, 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "") As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, sFilterForDataView, 1)
  End Function
  Public Function Insert() As Integer
    Dim iResult As Integer
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sb As New System.Text.StringBuilder()
    Dim sbv As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Insert = 0
    sb.Append("Insert into [CodeData](")
    If sInsUpdCodeKey.ToString <> "" Then
      sb.Append("CodeKey,")
      sbv.Append(sInsUpdCodeKey & ",")
    End If
    If sInsUpdCodeValue.ToString <> "" Then
      sb.Append("CodeValue,")
      sbv.Append(sInsUpdCodeValue & ",")
    End If
    If sInsUpdCodeValueLarge.ToString <> "" Then
      sb.Append("CodeValueLarge,")
      sbv.Append(sInsUpdCodeValueLarge & ",")
    End If

    sSQL = sb.ToString
    sSQL2 = sbv.ToString
    sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
    sSQL2 = Left(sSQL2, Len(sSQL2) - 1) & ")"

    sSQL = sSQL & sSQL2
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    iResult = oCmd.ExecuteNonQuery
    If iResult < 1 Then
      oUtil.CloseConnection(Connection, Transaction)
      Exit Function
    End If
    sSQL = "Select max(CodeID) from [CodeData]"
    oCmd.CommandText = sSQL
    Insert = oCmd.ExecuteScalar
    CodeID_PK__Integer = Insert.ToString
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    sbv = Nothing
    oCmd = Nothing
  End Function

  Public Function Update(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()

    Update = 0
    sb.Append("Update [CodeData] Set ")
    If sInsUpdCodeKey.ToString <> "" Then sb.Append("CodeKey=" & sInsUpdCodeKey & ",")
    If sInsUpdCodeValue.ToString <> "" Then sb.Append("CodeValue=" & sInsUpdCodeValue & ",")
    If sInsUpdCodeValueLarge.ToString <> "" Then sb.Append("CodeValueLarge=" & sInsUpdCodeValueLarge & ",")
    sSQL = sb.ToString
    If Right(sSQL, 1) = "," Then
      sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If sWhereClause <> Nothing Then
      sSQL = sSQL & " Where " & sWhereClause
    Else
      sSQL = sSQL & " Where CodeID=" & sInsUpdCodeID
    End If

    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    oCmd.CommandText = sSQL
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    Update = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Public Function Delete(Optional sWhereClause As String = "") As Integer
    Dim sSQL As String
    Dim sb As New System.Text.StringBuilder()
    Dim oCmd As New System.Data.SqlClient.SqlCommand()
    Delete = 0
    sb.Append("Delete [CodeData] Where ")
    If sWhereClause <> Nothing Then
      sb.Append(sWhereClause)
      sSQL = sb.ToString
    Else
      sb.Append("CodeID=" & sInsUpdCodeID)
      sSQL = sb.ToString
    End If
    oUtil.OpenConnection(Connection, Transaction, ConnectionString)
    oCmd.Connection = Connection
    If Not (Transaction Is Nothing) Then
      oCmd.Transaction = Transaction
    End If

    oCmd.CommandText = sSQL
    Delete = oCmd.ExecuteNonQuery
    oUtil.CloseConnection(Connection, Transaction)
    sb = Nothing
    oCmd = Nothing
  End Function

  Protected Overrides Sub Finalize()
    Transaction = Nothing
    Connection = Nothing
    SelectedData = Nothing
    CurrentRow = Nothing
    oUtil = Nothing
    MyBase.Finalize()
  End Sub

End Class
