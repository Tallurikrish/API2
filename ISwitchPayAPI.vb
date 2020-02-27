Imports DataClasses.DataClasses
Imports SwitchPayIntegration.Models

<ServiceContract>
Public Interface ISwitchPayAPI


    <OperationContract>
    Function ExecuteWorkflow(ApplicationID As Long, DestinationRepositoryName As String, WorkflowName As String, QueueName As String, PriorityName As String, Data As String) As String

    <OperationContract>
    Function GetURLForCreditProvider(ApplicationID As Long) As String

    <OperationContract>
    Function ActivateTerminal(MerchantRef As String, TerminalRef As String, OTP As String, FinancialInstitutionID As Long) As String

    <OperationContract>
    Function RegisterTerminal(MerchantRef As String, TerminalRef As String, FinancialInstitutionID As Long) As String

    <OperationContract>
    Sub UpdateApplicationReference(ApplicationID As Long, Reference As String)

    <OperationContract>
    Function CreatePMIApplication(MerchantID As Long, TerminalID As String, ApplicationRef As String, FinanceAmount As Double, IDNumber As String, MobileNumber As String, BankID As Integer, GenerateOTP As Boolean, FirstName As String, Surname As String, GrossIncome As Double, NettIncome As Double) As Response

    <OperationContract>
    Function ResendLastSMS(ID As Long) As String

    <OperationContract>
    Function SendSMS(ApplicationID As Long, Message As String) As String

    <OperationContract>
    Function GetApplicationData(ApplicationID As Long) As String

    <OperationContract>
    Function SendOTP(MerchantID As Long, ID As Long, OTPTypeID As SwitchPayAPI.OTPType) As Boolean

    <OperationContract>
    Function ReceiveOTP(MerchantID As Long, ID As Long, otpTypeId As SwitchPayAPI.OTPType, OTP As String, TryOthers As Boolean) As Boolean

    <OperationContract>
    Function GetMerchantData(MID As Long) As ApplicationData

    <OperationContract>
    Sub AcceptOffer(ApplicationID As Long)

    <OperationContract>
    Function ActionWFStep(ApplicationID As Long, Result As String, WFNames As List(Of String)) As Boolean

    <OperationContract>
    Function GetData(ID As Long, Type As Integer) As ApplicationData

    <OperationContract>
    Function grossapproved(ID As Long) As Double
    <OperationContract>
    Function ActiveDectivateMerchant(MerchantID As Long, Status As Boolean) As Response

    <OperationContract>
    Function AddTerminal(MerchantID As Long, ProductID As Long, TerminalID As String, MonthlyFee As Decimal, MerchantFee As Decimal, Term As Long, ActivationDate As Date) As Long

    <OperationContract>
    Function LoanzieCancel(ApplicationID As Long) As String

    <OperationContract>
    Function AddTerminals(dt As DataTable) As DataTable

    <OperationContract>
    Function AppCreation(MerchID As Long, MerchantRef As String, ApplicationRef As String, TerminalID As String, FinanceAmount As Double, IDNumber As String, MobileNumber As String, BankID As Integer, GenerateOTP As Boolean, FirstName As String, Surname As String, GrossIncome As Double, NettIncome As Double, DealType As String, Source As String) As Response

    <OperationContract>
    Function AppCreationInRepository(ApplicationID As Long, Environment As String, Repository As String) As String

    <OperationContract>
    Function CancelApplication(MerchantID As Long, TerminalID As String, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As Response

    <OperationContract>
    Function CheckBankDetails(ApplicationID) As String

    <OperationContract>
    Function CreateAdminWorkflow(ApplicationID As Long, QueueName As String, PriorityName As String) As String

    <OperationContract>
    Function CreateApplication(MerchantID As Long, TerminalID As String, ApplicationRef As String, FinanceAmount As Double, IDNumber As String, MobileNumber As String, BankID As Integer, GenerateOTP As Boolean, FirstName As String, Surname As String, GrossIncome As Double, NettIncome As Double) As Response
    <OperationContract>
    Function CreateApplicationTerminal(MerchantRef As String, TerminalID As String, ApplicationRef As String, FinanceAmount As Double, IDNumber As String, MobileNumber As String, BankID As Integer, GenerateOTP As Boolean, FirstName As String, Surname As String, GrossIncome As Double, NettIncome As Double) As Response

    <OperationContract>
    Function CreateApplicationWeb(MerchantID As Long, TerminalID As String, ApplicationRef As String, FinanceAmount As Double, IDNumber As String, MobileNumber As String, BankID As Integer, GenerateOTP As Boolean, FirstName As String, Surname As String, GrossIncome As Double, NettIncome As Double) As Response

    <OperationContract>
    Function AddMetric(MerchantID As Long, MerchantTerminalID As Long, MerchantRef As String, TerminalID As String, Key As String, Value As String) As Boolean

    <OperationContract>
    Function AddMetrics(MerchantID As Long, MerchantTerminalID As Long, MerchantRef As String, TerminalID As String, Pairs As Dictionary(Of String, String)) As Boolean

    <OperationContract>
    Sub CreateAuditItem(ApplicationID As Long, Details As String, AuditTypeID As Long, Optional SetStatus As Boolean = True)

    <OperationContract>
    Sub CreateAuditItemDetail(ApplicationID As Long, Details As String, AuditTypeID As Long, Name As String, IPAddress As String, Optional SetStatus As Boolean = True)

    <OperationContract>
    Function CreateDashboardWorkflow(ApplicationID As Long, QueueName As String, PriorityName As String) As String

    <OperationContract>
    Function CreateDPApplication(MerchantID As Long, TerminalID As String, ApplicationRef As String, FinanceAmount As Double, IDNumber As String, MobileNumber As String, BankID As Integer, GenerateOTP As Boolean) As Response

    <OperationContract>
    Function CreateLaybyApplication(MerchantID As Long, MerchantRef As String, TerminalID As String, ProductBandTermID As Long, ApplicationRef As String, FinanceAmount As Double, IDNumber As String, MobileNumber As String, BankID As Integer, GenerateOTP As Boolean, FirstName As String, Surname As String, GrossIncome As Double, NettIncome As Double, Term As Integer, Deposit As Double) As Response

    <OperationContract>
    Function CreateMerchant(Title As String, ContactName As String, BankAccount As String, Reference As String, Phone As String, Email As String, ShortName As String, RegisteredName As String, RegNo As String, ParentMerchantID As Long) As Long

    <OperationContract>
    Function CreateWorkflow(ApplicationID As Long, WorkflowName As String, QueueName As String, PriorityName As String) As String

    <OperationContract>
    Sub DealRejected(ApplicationID As Long)

    <OperationContract>
    Sub DeclineOffer(ApplicationID As Long)

    <OperationContract>
    Sub DeleteBankDetails(ApplicationID)

    <OperationContract>
    Sub GenerateCommInvoice(InvoiceID As Long)

    <OperationContract>
    Sub GenerateDebitOrders(MerchantID As Long)

    <OperationContract>
    Sub GenerateInvoice(InvoiceID As Long)

    <OperationContract>
    Sub GeneratePaymentRequest(ApplicationID As Long)

    <OperationContract>
    Function GenerateSubsFile(FileID As Long) As String

    <OperationContract>
    Function GetAdditionalFields(ApplicationID As Long, Responsestr As String) As String

    <OperationContract>
    Function GetApplicationFieldValue(ApplicationID As Long, FieldDefinitionEntityID As Long) As String

    <OperationContract>
    Function GetApplicationFieldValueCode(ApplicationID As Long, FieldDefinitionEntityID As Long) As String

    <OperationContract>
    Function GetApplications(MerchantID As Long, TerminalID As String, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As Response

    <OperationContract>
    Function GetApplicationStatus(MerchantID As Long, TerminalID As String, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As Response
    <OperationContract>
    Function GetApplicationStatusShort(ByVal ApplicationID As Long) As Response

    <OperationContract>
    Function GetBanksByMerchantID(MerchantID As Long, MerchantRef As String, TerminalID As String, Purpose As Integer) As DataTable

    <OperationContract>
    Function GetLastTransaction(MerchantID As Long, MerchantRef As String, TerminalID As String, Purpose As Integer) As Response

    <OperationContract>
    Function GetLayByDetails(MerchantID As Long, MerchantRef As String, TerminalID As String, Amount As Double, ProductBandTermID As Long, ProductID As Long) As Response

    <OperationContract>
    <WebInvoke>
    Function GetLookUps(MerchantID As Long, MerchantRef As String, TerminalID As String, Purpose As Integer) As LookUps

    <OperationContract>
    Function GetMerchantLogo(MerchantID As Long) As Byte()

    <OperationContract>
    Function GetMerchantName(MerchantID As Long) As String

    <OperationContract>
    Function GetMerchantStatus(Optional ByVal MerchantID As Long = 0, Optional ByVal MerchantRef As String = "", Optional ByVal TerminalID As String = "") As Response

    <OperationContract>
    Function GetTerms(MerchantID As Long, MerchantRef As String, TerminalID As String, Amount As Double, ProductID As Long) As Products

    <OperationContract>
    Function isActivity(ApplicationID As Long, DisplayName As String) As String

    <OperationContract>
    Function LoanzieAcceptOffer(ApplicationID As Long) As Decimal

    <OperationContract>
    Function LoanzieGetAgreement(ApplicationID As Long) As DataSet

    <OperationContract>
    Function LoanziePrevet(ApplicationID As Long) As String

    <OperationContract>
    Function CapitecPrevet(ApplicationID As Long) As String

    <OperationContract>
    Function LoanzieQuickCheck(ApplicationID As Long) As String

    <OperationContract>
    Function LoanzieQuickCheckOffer(ApplicationID As Long, AppID As String) As String

    <OperationContract>
    Sub LoanzieSaveData(ApplicationID As Long)

    <OperationContract>
    Function NIUSSDAuthorise(ApplicationID As Long, Response As String) As String

    <OperationContract>
    Sub NTUOffer(ApplicationID As Long)

    <OperationContract>
    Sub PaidFile(FileID As Long, Items As String)

    <OperationContract>
    Sub PaidCommFile(FileID As Long, Items As String)

    <OperationContract>
    Function Podium(Message As String) As Boolean

    <OperationContract>
    Function PreparePaymentRun(ApplicationID As Long)

    <OperationContract>
    Sub PreparePaymentsFile(FileID As Long)

    <OperationContract>
    Function ReceiveDeliveryReceipt(AppicationID As Long, Mobile As String, Message As String, Tag As String) As String

    <OperationContract>
    Function ReceiveNIUSSDMessage(AppicationID As Long, Mobile As String, Message As String, Tag As String) As String

    <OperationContract>
    Function ReceiveSMSMessage(AppicationID As Long, Mobile As String, Message As String, Tag As String) As String

    <OperationContract>
    Function ReceiveUSSDMessage(AppicationID As Long, Mobile As String, Message As String, Tag As String) As String

    <OperationContract>
    Function ReceiveVodacomDeliveryReceipt(AppicationID As Long, Mobile As String, Message As String, Tag As String) As String

    <OperationContract>
    Function ReceiveVodacomNIUSSDMessage(AppicationID As Long, Mobile As String, Message As String, Tag As String) As String

    <OperationContract>
    Function ReceiveVodacomSMSMessage(AppicationID As Long, Mobile As String, Message As String, Tag As String) As String


    <OperationContract>
    Function ReceiveVodacomUSSDMessage(AppicationID As Long, Mobile As String, Message As String, Tag As String) As String

    <OperationContract>
    Function RedeemApplication(MerchantID As Long, TerminalID As String, FinanceAmount As Double, GenerateOTP As Boolean, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As Response

    <OperationContract>
    Function RegisterMerchant(MerchantRef As String, Name As String, ContactNumber As String, URL As String, Email As String, IndustryID As Long) As Response

    <OperationContract>
    Function RegisterMerchantSkelta(MerchantRef As String, Name As String, ContactNumber As String, URL As String, Email As String, IndustryID As Long) As String



    <OperationContract>
    Sub ReleaseCommFile(FileID As Long)

    <OperationContract>
    Sub ReleaseFile(FileID As Long)

    <OperationContract>
    Function ReturnApplication(MerchantID As Long, TerminalID As String, FinanceAmount As Double, GenerateOTP As Boolean, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As Response

    <OperationContract>
    Sub SaveDisbursementFileItems(FileID As Long, Items As String)

    <OperationContract>
    Sub SaveDOFileItems(FileID As Long, Items As String)

    <OperationContract>
    Function SendApplicationOTP(MerchantID As Long, TerminalID As String, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As Response

    <OperationContract>
    Function SendAuthSMS(ApplicationID As Long) As String

    <OperationContract>
    Function SendBankMail(ApplicationID As Long, BankID As Long) As String

    <OperationContract>
    Function SendCollectedSMS(ApplicationID As Long) As String

    <OperationContract>
    Function SendCollectSMS(ApplicationID As Long) As String

    <OperationContract>
    Function SendCollectSMSOTP(ApplicationID As Long, OTP As String) As String

    <OperationContract>
    Function SendNIUSSD(ApplicationID As Long, Message As String) As String


    <OperationContract>
    Function SendOfferSMS(ApplicationID As Long) As String

    <OperationContract>
    Function SendRedeemAuthSMS(ApplicationID As Long) As String

    <OperationContract>
    Function SendRedeemOTP(MerchantID As Long, TerminalID As String, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As Response

    <OperationContract>
    Function SendReturnOTP(MerchantID As Long, TerminalID As String, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As Response


    <OperationContract>
    Function SendSMSToNumber(MobileNumber As String, Message As String) As String

    <OperationContract>
    Function SendSMSVodacom(ApplicationID As Long, Message As String) As String

    <OperationContract>
    Function SendVodacomNIUSSD(ApplicationID As Long, Message As String) As String

    <OperationContract>
    Function SendVodacomSMSURL(ApplicationID As Long, Message As String) As String

    <OperationContract>
    Function SubmitApplicationOTP(MerchantID As Long, TerminalID As String, OTP As String, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As Response

    <OperationContract>
    Function SubmitRedeemOTP(MerchantID As Long, TerminalID As String, OTP As String, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As Response

    <OperationContract>
    Function SubmitResult(ApplicationID As Long, BankID As Long, Successful As Boolean, OfferAmount As Double, OfferInstallment As Double, OfferTerm As Integer, Reference As String) As String

    <OperationContract>
    Function SubmitReturnOTP(MerchantID As Long, TerminalID As String, OTP As String, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As Response

    <OperationContract>
    Sub UnReleaseFile(FileID As Long)

    <OperationContract>
    Sub UnReleaseCommFile(FileID As Long)

    <OperationContract>
    Function UpdateApplicationStatus(MerchantID As Long, TerminalID As String, NewStatusID As Integer, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As Response

    <OperationContract>
    Sub UpdateDebitOrders()
    <OperationContract>
    Function ValidateID(ByVal IDno As String) As Boolean

End Interface
