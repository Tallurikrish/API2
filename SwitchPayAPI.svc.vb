' NOTE: You can use the "Rename" command on the context menu to change the class name "Service1" in code, svc and config file together.
Imports DataClasses.DataClasses
#Disable Warning BC40056 ' Namespace or type specified in the Imports 'Microsoft.ApplicationInsights' doesn't contain any public member or cannot be found. Make sure the namespace or the type is defined and contains at least one public member. Make sure the imported element name doesn't use any aliases.
Imports Microsoft.ApplicationInsights
#Enable Warning BC40056 ' Namespace or type specified in the Imports 'Microsoft.ApplicationInsights' doesn't contain any public member or cannot be found. Make sure the namespace or the type is defined and contains at least one public member. Make sure the imported element name doesn't use any aliases.
#Disable Warning BC40056 ' Namespace or type specified in the Imports 'Microsoft.ApplicationInsights.Wcf' doesn't contain any public member or cannot be found. Make sure the namespace or the type is defined and contains at least one public member. Make sure the imported element name doesn't use any aliases.
Imports Microsoft.ApplicationInsights.Wcf
#Enable Warning BC40056 ' Namespace or type specified in the Imports 'Microsoft.ApplicationInsights.Wcf' doesn't contain any public member or cannot be found. Make sure the namespace or the type is defined and contains at least one public member. Make sure the imported element name doesn't use any aliases.
' NOTE: In order to launch WCF Test Client for testing this service, please select Service1.svc or Service1.svc.vb at the Solution Explorer and start debugging.
Imports System.Net
Imports System.Net.Mail
Imports FastSerialization
Imports SwitchPayIntegration.Models
Imports System.Data.Entity.Core.Common.CommandTrees
Imports API2.My
Imports Microsoft.Diagnostics.Tracing.Parsers.IIS_Trace
Imports System.IO

'<ServiceTelemetry>
Public Class SwitchPayAPI
    Implements ISwitchPayAPI

    Private MRef As String = ""
    Private TtRef As String = ""

    Enum OTPType
        MerchantCreation = 1
        MerchantUpdating = 2
        MerchantActivation = 3
        TerminalDeActivation = 4
        TerminalActivation = 5
        ApplicationCreation = 6
        ApplicationCollection = 7
    End Enum

    Function CreatePMIApplication(MerchantID As Long, TerminalID As String, ApplicationRef As String, FinanceAmount As Double, IDNumber As String, MobileNumber As String, BankID As Integer, GenerateOTP As Boolean, FirstName As String, Surname As String, GrossIncome As Double, NettIncome As Double) As Response Implements ISwitchPayAPI.CreatePMIApplication
        Dim x = AppCreation(MerchantID, String.Empty, ApplicationRef, TerminalID, FinanceAmount, IDNumber, MobileNumber, BankID, GenerateOTP, FirstName, Surname, GrossIncome, NettIncome, "PMI", "Integration")
    End Function

    Function ResendLastSMS(ID As Long) As String Implements ISwitchPayAPI.ResendLastSMS

        Return SendSMS(ID, "Resending previous SMS")
    End Function

    Function GetApplicationData(ApplicationID As Long) As String Implements ISwitchPayAPI.GetApplicationData
        'Dim appdata As New Microsoft.Bot.Schema.Activity
        'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        'Dim aus = (From q In db.vHistories Where q.ApplicationID = ApplicationID Select q).ToList
        'For Each au In aus
        '    appdata.Attachments.Add(New Microsoft.Bot.Schema.Attachment(, , au.AuditDate & " - " & au.Type & ": " & au.Details))
        'Next
        Return ""
    End Function

    Function CanCollect(ApplicationID As Long) As Boolean
        Try
            Dim dd As New DataDictionary(ApplicationID)
            Return dd.ApplicationD.CanCollect()
        Catch ex As Exception
            Return False
        End Try
    End Function

    Function RegisterTerminal(MerchantRef As String, TerminalRef As String, FinancialInstitutionID As Long) As String Implements ISwitchPayAPI.RegisterTerminal
        Try
            dd = New DataDictionary("NB", MerchantRef, "Terminal")
        Catch
            Return "Problem Loading Merchant/Terminal"
        End Try
        If dd.AppMerchantTerminal Is Nothing Then
            Return "No Terminal Found Ready For Activation"
        End If
        MRef = MerchantRef
        TtRef = TerminalRef
        Try
            If SendOTP(dd.AppMerchantTerminal.MerchantID, dd.AppMerchantTerminal.ID, 5) Then
                Return "OTP Sent Successfully" & dd.AppMerchant.ID
            Else
                Return "OTP Failed To Send" & dd.AppMerchant.ID
            End If
        Catch ex As Exception
            Return ex.Message
        End Try

    End Function

    Function ActivateTerminal(MerchantRef As String, TerminalRef As String, OTP As String, FinancialInstitutionID As Long) As String Implements ISwitchPayAPI.ActivateTerminal
        Try
            dd = New DataDictionary("NB", MerchantRef, "Terminal")
        Catch ex As Exception
            Return ex.Message
            'Return "Problem Loading Merchant/Terminal"
        End Try
        If dd.AppMerchantTerminal Is Nothing Then
            Return "No Terminal Found Waiting For OTP"
        End If
        MRef = MerchantRef
        TtRef = TerminalRef
        Try
            'If ReceiveOTP(dd.AppMerchantTerminal.ID, 5, OTP, False) Then
            If ReceiveOTP(dd.AppMerchantTerminal.MerchantID, dd.AppMerchantTerminal.ID, 5, OTP, False) Then
                Return "Terminal Successfully Activated" & dd.AppMerchant.ID
            Else
                Return "Terminal Failed To Activate" & dd.AppMerchant.ID
            End If
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Function ExecuteWorkflow(ApplicationID As Long, DestinationRepositoryName As String, WorkflowName As String, QueueName As String, PriorityName As String, Data As String) As String Implements ISwitchPayAPI.ExecuteWorkflow
        Dim DataD As New DataDictionary(ApplicationID, My.Settings.Environment, IIf(DestinationRepositoryName = "", My.Settings.Repository, DestinationRepositoryName), My.Settings.Environment, My.Settings.Repository)
        Return DataD.WorkflowD.ExecuteWorkflow(WorkflowName, QueueName, PriorityName)
    End Function

    Function SendOTP(MerchantID As Long, ID As Long, otpTypeId As SwitchPayAPI.OTPType) As Boolean Implements ISwitchPayAPI.SendOTP
        Dim dd As New DataDictionary
        Dim bresult As Boolean = False
        Dim result As Boolean = False
        Select Case otpTypeId
            Case OTPType.ApplicationCollection
                dd = New DataDictionary(ID)
                result = dd.ApplicationD.SendApplyOTP()
            Case OTPType.ApplicationCreation
                dd = New DataDictionary(ID)
                result = dd.ApplicationD.SendCollectOTP()
            Case OTPType.MerchantActivation
                dd.LoadMerchantByID(ID)
                result = dd.MerchantD.SendMerchantActivationOTP()
            Case OTPType.MerchantCreation
                dd.LoadMerchantByID(ID)
                result = dd.MerchantD.SendMerchantCreationOTP()
            Case OTPType.MerchantUpdating
                dd.LoadMerchantByID(ID)
                result = dd.MerchantD.SendMerchantUpdatingOTP()
            Case OTPType.TerminalActivation
                dd.LoadMerchantByID(MerchantID, "", ID)

                result = dd.MerchantD.SendTerminalActivationOTP()
            Case OTPType.TerminalDeActivation
                dd.LoadMerchantByID(MerchantID, "", ID)
                result = dd.MerchantD.SendTerminalDeActivationOTP()
            Case Else
                result = False
        End Select
        Return result
    End Function

    Function ReceiveOTP(MerchantID As Long, ID As Long, otpTypeId As SwitchPayAPI.OTPType, OTP As String, TryOthers As Boolean) As Boolean Implements ISwitchPayAPI.ReceiveOTP
        Dim dd As DataDictionary
        Dim bresult As Boolean = False
        Select Case otpTypeId
            Case OTPType.ApplicationCollection
                dd = New DataDictionary(ID)
                bresult = dd.ApplicationD.ReceiveCollectOTP(OTP)
            Case OTPType.ApplicationCreation
                dd = New DataDictionary(ID)
                bresult = dd.ApplicationD.ReceiveApplyOTP(OTP)
            Case OTPType.MerchantActivation
                dd = New DataDictionary()
                dd.LoadMerchantByID(ID)
                bresult = dd.MerchantD.ReceiveMerchantOTP(otpTypeId, OTP, False)
            Case OTPType.MerchantCreation
                dd = New DataDictionary()
                dd.LoadMerchantByID(ID)
                bresult = dd.MerchantD.ReceiveMerchantOTP(otpTypeId, OTP, False)
            Case OTPType.MerchantUpdating
                dd = New DataDictionary()
                dd.LoadMerchantByID(ID)
                bresult = dd.MerchantD.ReceiveMerchantOTP(otpTypeId, OTP, False)
            Case OTPType.TerminalActivation
                dd = New DataDictionary()
                dd.LoadMerchantByID(MerchantID, "", ID)
                dd.AppMerchantTerminal.Reference = TtRef
                bresult = dd.MerchantD.ReceiveTerminalOTP(otpTypeId, OTP, False)
            Case OTPType.TerminalDeActivation
                dd = New DataDictionary()
                dd.LoadMerchantByID(0, "", ID)
                bresult = dd.MerchantD.ReceiveTerminalOTP(otpTypeId, OTP, False)
            Case Else
                bresult = False
        End Select
        Return bresult


    End Function

    Function GetMerchantData(MID As Long) As ApplicationData Implements ISwitchPayAPI.GetMerchantData
        Dim merdata As New ApplicationData
        Return merdata
    End Function

    Public Sub AcceptOffer(ApplicationID As Long) Implements ISwitchPayAPI.AcceptOffer
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim a
        Try
            a = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
            Try
                Dim wfs As New List(Of String)
                wfs.Add("Loanzie")
                ActionWFStep(ApplicationID, "Approved", wfs)

            Catch ex As Exception

            End Try
            Dim au As New Audit
            au.Name = "System"
            au.ApplicationID = a.ID
            au.AuditDate = Now
            a.AuditTypeID = 12
            au.Details = "Client Contracted"
            au.AuditTypeID = 12
            a.AuditTypeID = 12
            db.Audits.InsertOnSubmit(au)
            db.SubmitChanges()
            SendCollectSMS(ApplicationID)
            Dim email2 As New MailMessage


            Dim SMTP As New SmtpClient("smtp.gmail.com")

            email2.From = New MailAddress("workflow@switchpay.co.za")
            SMTP.UseDefaultCredentials = False
            SMTP.Credentials = New System.Net.NetworkCredential("workflow@switchpay.co.za", "selfadrpcbiajyux") '<-- Password Here
            SMTP.EnableSsl = True
            email2.Subject = a.Reference & " - Contracted"
            email2.To.Add("hendrik@acpas.co.za")
            email2.To.Add("jaco@acpas.co.za")
            email2.To.Add("support@acpas.co.za")
            email2.To.Add("diani@ammacom.com")
            email2.To.Add("Sacha.Craig@pmi.com")
            email2.To.Add("iqos@loanzie.co.za")

            email2.IsBodyHtml = True
            email2.Body = "Client Contracted Deal<br />"
            Dim hist = (From q In db.vHistories Where q.ApplicationID = CLng(ApplicationID) Select q).ToArray()
            For Each h In hist
                email2.Body = email2.Body & h.AuditDate.ToString() & "<br />" & h.Details & "<br />"
            Next
            SMTP.Port = "587"
            SMTP.Send(email2)
            Try
                Dim wfs As New List(Of String)
                wfs.Add("Loanzie")
                ActionWFStep(ApplicationID, "Approved", wfs)

            Catch ex As Exception

            End Try
        Catch ex As Exception
            Try
                Dim wfs As New List(Of String)
                wfs.Add("Loanzie")
                ActionWFStep(ApplicationID, "Approved", wfs)

            Catch ss As Exception

            End Try
            Dim au As New Audit
            au.Name = "System"
            au.ApplicationID = a.ID
            au.AuditDate = Now
            au.Details = "Deal Contracted"
            a.AuditTypeID = 12
            au.AuditTypeID = 12
            db.SubmitChanges()
            db.Audits.InsertOnSubmit(au)
            db.SubmitChanges()
            SendCollectSMS(a.ID)
        End Try
    End Sub

    Public Function ActionWFStepRepository(ApplicationID As Long, Result As String, WFNames As List(Of String), Repository As String) As Boolean
        Dim bresult As Boolean = False
        Try
            Dim dd As New DataDictionary(ApplicationID, My.Settings.Environment, Repository, My.Settings.Environment, Repository)
            Dim Success = dd.WorkflowD.ActionWorkItem(Result, WFNames)
        Catch ex As Exception
            Throw ex
        End Try
        Return bresult
    End Function

    Public Function ActionWFStep(ApplicationID As Long, Result As String, WFNames As List(Of String)) As Boolean Implements ISwitchPayAPI.ActionWFStep
        Dim bresult As Boolean = False
        Try
            Dim dd As New DataDictionary(ApplicationID, My.Settings.Environment, My.Settings.Repository, My.Settings.Environment, My.Settings.Repository)
            Dim Success = dd.WorkflowD.ActionWorkItem(Result, WFNames)
        Catch ex As Exception
            Throw ex
        End Try
        Return bresult
    End Function

    Function ActiveDectivateMerchant(MerchantID As Long, Status As Boolean) As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.ActiveDectivateMerchant
        Dim rxsd As New SwitchPayIntegration.Models.Response
        Dim hasMerchant As Boolean = False
        Try
            Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            hdr = hdt.NewHeaderRow()
            hdr("MerchantID") = MerchantID
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Merchant = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
            hasMerchant = True
            hdr("MerchantRef") = Merchant.Reference
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim fdr As SwitchPayIntegration.Models.Response.FieldsRow
            fdr = fdt.NewFieldsRow()
            fdr("Name") = "StatusID"
            fdr("Value") = 2
            fdt.AddFieldsRow(fdr)
            Dim fdr2 As SwitchPayIntegration.Models.Response.FieldsRow
            fdr2 = fdt.NewFieldsRow()
            fdr2("Name") = "Status"
            fdr2("Value") = "Active"
            fdt.AddFieldsRow(fdr2)
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "0"
            mdr("Message") = "Merchant ID:  " & MerchantID & " status successfully checked"
            mdr("IsError") = False
            mdt.AddMessagesRow(mdr)

        Catch ex As Exception
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            'Dim apps = (From q In db.Applications Where appl)
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            hdr = hdt.NewHeaderRow()
            hdr("IsError") = True
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If hasMerchant = False Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "4"
                mdr3("Message") = "No Valid Merchant Found."
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "1"
            mdr("Message") = "Exception: " & ex.Message
            mdr("IsError") = True
            mdt.AddMessagesRow(mdr)
        End Try
        Return rxsd
    End Function

    Public Function GetData(ID As Long, Type As Integer) As ApplicationData Implements ISwitchPayAPI.GetData
        Dim dict As New DataDictionary()
        Return New ApplicationData()
    End Function


    Function AddTerminal(MerchantID As Long, ProductID As Long, TerminalID As String, MonthlyFee As Decimal, MerchantFee As Decimal, Term As Long, ActivationDate As Date) As Long Implements ISwitchPayAPI.AddTerminal
        If ProductID = 1 Then
            Return 2
        Else
            Throw New Exception("There wan an error")
        End If
        Return 1
    End Function

    Function AddTerminals(dt As DataTable) As DataTable Implements ISwitchPayAPI.AddTerminals

        Return New DataTable("Terminals")
    End Function

    Public Function AppCreation(MerchID As Long, MerchantRef As String, ApplicationRef As String, TerminalID As String, FinanceAmount As Double, IDNumber As String, MobileNumber As String, BankID As Integer, GenerateOTP As Boolean, FirstName As String, Surname As String, GrossIncome As Double, NettIncome As Double, DealType As String, Source As String) As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.AppCreation
        Dim rxsd As New SwitchPayIntegration.Models.Response
        Dim result As String = String.Empty
        Dim s As String = "0"

        Try
            If MerchantRef = String.Empty Then
                dd = New DataDictionary(TerminalID, MerchID, Source)
            Else
                dd = New DataDictionary(TerminalID, MerchantRef, Source)
            End If
            Try
                s &= "1Merchant: " & dd.AppMerchant.ID & " Terminal: " & dd.AppMerchantTerminal.ID.ToString & vbCrLf
            Catch
            End Try
            Try
                s &= "37Merchant: " & dd.ApplicationD.App.MerchantID & " Terminal: " & dd.ApplicationD.App.MerchantTerminalID.ToString & vbCrLf
            Catch
            End Try
            rxsd = dd.AppCreation(MerchID, MerchantRef, ApplicationRef, TerminalID, FinanceAmount, IDNumber, MobileNumber, BankID, False, FirstName, Surname, GrossIncome, NettIncome, DealType, Source)
            Try
                s &= "2Merchant: " & dd.AppMerchant.ID & " Terminal: " & dd.AppMerchantTerminal.ID.ToString & vbCrLf
            Catch
            End Try
            Try
                s &= "2ss7Merchant: " & dd.ApplicationD.App.MerchantID & " Terminal: " & dd.ApplicationD.App.MerchantTerminalID.ToString & vbCrLf
            Catch
            End Try
            result = "Success"
        Catch ex As Exception
            result = ex.Message
            Try
                result = result & " Inner: " & ex.InnerException.Message
            Catch
            End Try
        End Try
        If result <> "Success" Then
            Throw New Exception(result)
        End If
        If (rxsd.Header.Rows(0)("IsError") = True) And IsNumeric(rxsd.Header.Rows(0)("ApplicationID")) Then
            dd.ApplicationD.CreateAuditItem("Deal Void - Couldnt start workflow", 66, "System")
        Else
            Try
                If rxsd.Header.Rows(0)("IsError") = False Then
                    Try
                        s &= "3Merchant: " & dd.AppMerchant.ID & " Terminal: " & dd.AppMerchantTerminal.ID.ToString & vbCrLf
                    Catch
                    End Try
                    dd.PopulateMainDBObjects()
                    Try
                        s &= "4Merchant: " & dd.AppMerchant.ID & " Terminal: " & dd.AppMerchantTerminal.ID.ToString & vbCrLf
                    Catch
                    End Try
                    dd.PopulateChildDBObjects()
                    Try
                        s &= "5Merchant: " & dd.AppMerchant.ID & " Terminal: " & dd.AppMerchantTerminal.ID.ToString & vbCrLf
                    Catch
                    End Try
                    dd.PopulateREObjects()
                    Try
                        s &= "6Merchant: " & dd.AppMerchant.ID & " Terminal: " & dd.AppMerchantTerminal.ID.ToString & vbCrLf
                    Catch
                    End Try
                    Try
                        s &= "7Merchant: " & dd.ApplicationD.App.MerchantID & " Terminal: " & dd.ApplicationD.App.MerchantTerminalID.ToString & vbCrLf
                    Catch
                    End Try

                    If DealType = "PMI" Then
                        result = dd.WorkflowD.ExecuteWorkflow("New PMI Application", "Deals", "Low")
                    Else
                        result = dd.WorkflowD.ExecuteWorkflow("New Application", "Deals", "Low")
                    End If
                    'If IDNumber = "1111111111111" Then
                    '    AutoDeal(rxsd)
                    'End If
                End If
            Catch ex As Exception
                result = "Complete Load - " & ex.Message
                Try
                    result = result & " Inner: " & ex.InnerException.Message
                Catch
                End Try
            End Try


        End If
        If result <> "Success" Then
            rxsd.Header.Rows(0)("IsError") = True
            Dim mdt As SwitchPayIntegration.Models.Response.MessagesDataTable = rxsd.Tables("Messages")
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "5"
            mdr("Message") = "Application ID: " & rxsd.Header.Rows(0)("ApplicationID") & " failed with error: " & result
            mdr("IsError") = True
            mdt.AddMessagesRow(mdr)
            rxsd.AcceptChanges()
        End If
        Return rxsd
    End Function

    Sub AutoDeal(a As Response)
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim am As Long
        Dim otp As String
        Dim w As Application
        Threading.Thread.Sleep(2000)
        am = a.Header.Rows(0)("ApplicationID")

        w = (From q In db.Applications Where q.ID = am Select q Order By q.ID Descending).First()
        otp = w.AcceptPIN
        w.OfferAmount = w.FinanceAmount

        db.SubmitChanges()
        a = SubmitApplicationOTP(w.MerchantID, w.MerchantTerminalID, w.AcceptPIN, am, w.Reference, w.MobileNumber, w.IDNumber)
        Dim wfs As New List(Of String)
        wfs.Add("Additional Fields")
        wfs.Add("Loanzie")
        Try
            Try
                ActionWFStep(am, "Approved", wfs)
            Catch ex3 As Exception
            End Try
        Catch
        End Try
        Try
            ActionWFStep(am, "Approved", wfs)
        Catch ex3 As Exception
        End Try
        RedeemApplication(w.MerchantID, w.MerchantTerminalID, w.OfferAmount, False, w.ID, w.Reference, w.MobileNumber, w.IDNumber)
        am = a.Header.Rows(0)("ApplicationID")
        w = (From q In db.Applications Where q.ID = am Select q Order By q.ID Descending).First()
        otp = w.CollectPIN
        a = SubmitRedeemOTP(w.MerchantID, w.MerchantTerminalID, w.CollectPIN, w.ID, w.Reference, w.MobileNumber, w.IDNumber)
    End Sub

    Public Function AppCreationInRepository(ApplicationID As Long, EnvironmentName As String, RepositoryName As String) As String Implements ISwitchPayAPI.AppCreationInRepository
        Try
            Dim dd As New DataDictionary(CLng(ApplicationID), EnvironmentName, CStr(IIf(EnvironmentName = "Production", "", EnvironmentName) & RepositoryName), EnvironmentName, EnvironmentName)
            Dim result
            result = dd.WorkflowD.ExecuteCreationWorkflow(CStr(EnvironmentName), CStr(IIf(EnvironmentName = "Production", "", EnvironmentName) & RepositoryName))
            If result <> "Success" Then
                Throw New Exception(result)
            End If
            Return result
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Function CancelApplication(MerchantID As Long, TerminalID As String, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.CancelApplication
        Dim rxsd As New SwitchPayIntegration.Models.Response
        Dim hasMerchant As Boolean = False
        Dim hasIDNumber As Boolean = False
        Dim hasMobile As Boolean = False
        Dim nofilters As Boolean = False
        Dim results As Boolean = True
        Dim amountvalid As Boolean = False
        Dim bankvalid As Boolean = False
        Dim mobilevalid As Boolean = True
        Dim idvalid As Boolean = True
        Dim alreadycancelled As Boolean = False
        If (Not IsNumeric(IDNumber)) And (IDNumber <> String.Empty) Then
            idvalid = False
        End If
        If (Not IsNumeric(MobileNumber)) And (MobileNumber <> String.Empty) Then
            mobilevalid = False
        End If
        Try
            Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
            If (ApplicationID = 0) And (ApplicationRef = String.Empty) And (IDNumber = String.Empty) And (MobileNumber = String.Empty) Then
                nofilters = True
            End If
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Merchant = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
            hasMerchant = True

            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            Try
                If ApplicationID <> 0 Then
                    App = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
                ElseIf ApplicationRef <> String.Empty Then
                    App = (From q In db.Applications Where q.Reference = CStr(ApplicationRef) Select q Order By q.DateCreated Descending).First()
                ElseIf IDNumber <> String.Empty Then
                    App = (From q In db.Applications Where q.IDNumber = CStr(IDNumber) Select q Order By q.DateCreated Descending).First()
                Else
                    'Apps = (From q In db.Applications Where q.entityi = CStr(ApplicationRef) Select q).ToArray
                    results = False
                End If
            Catch
                results = False
            End Try
            hdr = hdt.NewHeaderRow()
            'LoadApplicationData(App.ID)
            hdr("MerchantID") = MerchantID
            hdr("MerchantRef") = Merchant.Reference
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If results Then
                If App.InternalAuditTypeID = 4 Then
                    alreadycancelled = True
                End If
                If alreadycancelled Then
                    hdr("ApplicationID") = App.ID
                    hdr("ApplicationRef") = App.Reference
                    hdr("IDNumber") = App.IDNumber
                    hdr("MobileNumber") = App.MobileNumber
                    hdr("IsError") = True
                    Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                    mdr2 = mdt.NewMessagesRow()
                    mdr2("Code") = "5"
                    mdr2("Message") = "Application ID: " & App.ID & " already cancelled"
                    mdr2("IsError") = True
                    mdt.AddMessagesRow(mdr2)
                Else
                    hdr("IsError") = False
                    App.InternalAuditTypeID = 4
                    db.SubmitChanges()
                    Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                    mdr2 = mdt.NewMessagesRow()
                    mdr2("Code") = "0"
                    mdr2("Message") = "Application ID: " & App.ID & " successfully cancelled"
                    mdr2("IsError") = False
                    mdt.AddMessagesRow(mdr2)

                End If
            Else

                hdr("IsError") = True
                If Not results Then
                    Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                    mdr2 = mdt.NewMessagesRow()
                    mdr2("Code") = "2"
                    mdr2("Message") = "No Applications Found."
                    mdr2("IsError") = True
                    mdt.AddMessagesRow(mdr2)
                End If
                If nofilters Then
                    Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                    mdr2 = mdt.NewMessagesRow()
                    mdr2("Code") = "3"
                    mdr2("Message") = "No Filters Provided."
                    mdr2("IsError") = True
                    mdt.AddMessagesRow(mdr2)
                End If
                If Not mobilevalid Then
                    Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                    mdr3 = mdt.NewMessagesRow()
                    mdr3("Code") = "5"
                    mdr3("Message") = "Mobile Number Invalid"
                    mdr3("IsError") = True
                    mdt.AddMessagesRow(mdr3)

                End If

                If Not idvalid Then
                    Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                    mdr3 = mdt.NewMessagesRow()
                    mdr3("Code") = "5"
                    mdr3("Message") = "ID Number Invalid"
                    mdr3("IsError") = True
                    mdt.AddMessagesRow(mdr3)


                End If
            End If

            hdt.AddHeaderRow(hdr)

        Catch ex As Exception
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            'Dim apps = (From q In db.Applications Where appl)
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            hdr = hdt.NewHeaderRow()
            hdr("IsError") = True
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim fdr2 As SwitchPayIntegration.Models.Response.FieldsRow
            fdr2 = fdt.NewFieldsRow()
            fdr2("Name") = "Screen Message"
            fdr2("Value") = "Cancelled"
            fdt.AddFieldsRow(fdr2)
            Dim fdr3 As SwitchPayIntegration.Models.Response.FieldsRow
            fdr3 = fdt.NewFieldsRow()
            fdr3("Name") = "SlipMessage"
            fdr3("Value") = "Application Cancelled"
            fdt.AddFieldsRow(fdr3)
            Dim fdr4 As SwitchPayIntegration.Models.Response.FieldsRow
            fdr4 = fdt.NewFieldsRow()
            fdr4("Name") = "TransactionType"
            fdr4("Value") = "Cancellation"
            fdt.AddFieldsRow(fdr4)
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")

            If hasMerchant = False Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "4"
                mdr3("Message") = "No Valid Merchant Found."
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "1"
            mdr("Message") = "Exception: " & ex.Message
            mdr("IsError") = True
            mdt.AddMessagesRow(mdr)
        End Try
        Return rxsd
    End Function
    Public Function CheckBankDetails(ApplicationID) As String Implements ISwitchPayAPI.CheckBankDetails
        Try
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Dim app = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
            Dim avs As New CompuscanAVSR.AVSRMain
            Dim acc As New CompuscanAVSR.Account
            acc.AccountNo = GetApplicationFieldValue(ApplicationID, 249)
            acc.IDNumber = app.IDNumber
            acc.Initials = app.FirstName.Substring(0, 1)
            acc.Surname = app.Surname
            Select Case GetApplicationFieldValueCode(ApplicationID, 248)
                Case 1
                    acc.BranchCode = "632005"
                Case 4
                    acc.BranchCode = "198765"
                Case 2
                    acc.BranchCode = "051001"

                Case 6
                    acc.BranchCode = "470010"

                Case 3
                    acc.BranchCode = "250655"

                Case 15
                    acc.BranchCode = "580105"
            End Select

            Select Case GetApplicationFieldValueCode(ApplicationID, 250)
                Case 23
                    acc.AccountType = CompuscanAVSR.AccountTypeEnum.Current
                Case 24
                    acc.AccountType = CompuscanAVSR.AccountTypeEnum.Savings
                Case 25
                    acc.AccountType = CompuscanAVSR.AccountTypeEnum.CreditCard
                Case 26
                    acc.AccountType = CompuscanAVSR.AccountTypeEnum.Bond
            End Select
            Dim x = avs.Validate(acc)
            Try
                CreateAuditItem(ApplicationID, "Code: " & x.Code & " - Message: " & x.Message & " - Xml: " & x.XML.ToString(), 62, True)
            Catch
            End Try
            Return x.Code.ToString()
        Catch ex As Exception
            Try
                CreateAuditItem(ApplicationID, "Error: " & ex.Message, 62, True)
            Catch
            End Try
            Return "3"
        End Try
    End Function

    Public Function CreateAdminWorkflow(ApplicationID As Long, QueueName As String, PriorityName As String) As String Implements ISwitchPayAPI.CreateAdminWorkflow
        Try
            Dim dd As New DataDictionary(ApplicationID)
            Dim result = dd.WorkflowD.ExecuteAdminWorkflow(QueueName, PriorityName)
            Return result
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try

    End Function

    Function CreateApplication(MerchantID As Long, TerminalID As String, ApplicationRef As String, FinanceAmount As Double, IDNumber As String, MobileNumber As String, BankID As Integer, GenerateOTP As Boolean, FirstName As String, Surname As String, GrossIncome As Double, NettIncome As Double) As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.CreateApplication
        Dim x = AppCreation(MerchantID, String.Empty, ApplicationRef, TerminalID, FinanceAmount, IDNumber, MobileNumber, BankID, GenerateOTP, FirstName, Surname, GrossIncome, NettIncome, "PBL", "Terminal")

        Return x
    End Function

    Function CreateApplicationTerminal(MerchantRef As String, TerminalID As String, ApplicationRef As String, FinanceAmount As Double, IDNumber As String, MobileNumber As String, BankID As Integer, GenerateOTP As Boolean, FirstName As String, Surname As String, GrossIncome As Double, NettIncome As Double) As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.CreateApplicationTerminal
        Return AppCreation(0, MerchantRef, ApplicationRef, TerminalID, FinanceAmount, IDNumber, MobileNumber, BankID, GenerateOTP, FirstName, Surname, GrossIncome, NettIncome, "PBL", "Terminal")
    End Function



    Function CreateApplicationWeb(MerchantID As Long, TerminalID As String, ApplicationRef As String, FinanceAmount As Double, IDNumber As String, MobileNumber As String, BankID As Integer, GenerateOTP As Boolean, FirstName As String, Surname As String, GrossIncome As Double, NettIncome As Double) As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.CreateApplicationWeb
        Return AppCreation(MerchantID, String.Empty, ApplicationRef, TerminalID, FinanceAmount, IDNumber, MobileNumber, BankID, False, FirstName, Surname, GrossIncome, NettIncome, "PBL", "Web")
    End Function

    Public Function AddMetric(MerchantID As Long, MerchantTerminalID As Long, MerchantRef As String, TerminalID As String, Key As String, Value As String) As Boolean Implements ISwitchPayAPI.AddMetric
        Try
            Dim dd As New DataDictionary(My.Settings.Environment, My.Settings.Repository)

            If MerchantRef = "" And TerminalID = "" Then
                dd.LoadMerchantByID(MerchantID, TerminalID, MerchantTerminalID)
            ElseIf MerchantRef = "" Then
                dd.LoadMerchantByID(MerchantID, TerminalID, MerchantTerminalID)
            ElseIf TerminalID = "" Then
                dd.LoadMerchantByReference(MerchantRef, TerminalID, MerchantTerminalID)
            End If

            Return dd.MerchantD.AddMetric(Key, Value)

        Catch
            Return False
        End Try
    End Function
    Public Function AddMetrics(MerchantID As Long, MerchantTerminalID As Long, MerchantRef As String, TerminalID As String, Pairs As Dictionary(Of String, String)) As Boolean Implements ISwitchPayAPI.AddMetrics
        Try
            Dim dd As New DataDictionary(My.Settings.Environment, My.Settings.Repository)

            If MerchantRef = "" And TerminalID = "" Then
                dd.LoadMerchantByID(MerchantID, TerminalID, MerchantTerminalID)
            ElseIf MerchantRef = "" Then
                dd.LoadMerchantByID(MerchantID, TerminalID, MerchantTerminalID)
            ElseIf TerminalID = "" Then
                dd.LoadMerchantByReference(MerchantRef, TerminalID, MerchantTerminalID)
            End If

            Return dd.MerchantD.AddMetrics(Pairs)

        Catch
            Return False
        End Try

    End Function

    Sub CreateAuditItem(ApplicationID As Long, Details As String, AuditTypeID As Long, Optional SetStatus As Boolean = True) Implements ISwitchPayAPI.CreateAuditItem
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim au As New Audit
        au.Details = Details
        au.Name = "System"
        au.AuditDate = Now
        au.ApplicationID = ApplicationID
        au.AuditTypeID = AuditTypeID
        db.Audits.InsertOnSubmit(au)
        Try
            If SetStatus Then
                App.AuditTypeID = AuditTypeID
            End If
        Catch
        End Try
        db.SubmitChanges()
    End Sub

    Sub CreateAuditItemDetail(ApplicationID As Long, Details As String, AuditTypeID As Long, Name As String, IPAddress As String, Optional SetStatus As Boolean = True) Implements ISwitchPayAPI.CreateAuditItemDetail
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim au As New Audit
        au.Details = Details
        au.Name = "System"
        au.Name = Name
        au.IPAddress = IPAddress
        au.AuditDate = Now
        au.ApplicationID = ApplicationID
        au.AuditTypeID = AuditTypeID
        db.Audits.InsertOnSubmit(au)
        Try
            If SetStatus Then
                App.AuditTypeID = AuditTypeID
            End If
        Catch
        End Try
        db.SubmitChanges()
    End Sub

    Public Function CreateDashboardWorkflow(ApplicationID As Long, QueueName As String, PriorityName As String) As String Implements ISwitchPayAPI.CreateDashboardWorkflow
        Try
            Dim dd As New DataDictionary(ApplicationID)
            Dim result = dd.WorkflowD.ExecuteDashboardWorkflow(QueueName, PriorityName)
            Return result
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try

    End Function

    Function CreateDPApplication(MerchantID As Long, TerminalID As String, ApplicationRef As String, FinanceAmount As Double, IDNumber As String, MobileNumber As String, BankID As Integer, GenerateOTP As Boolean) As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.CreateDPApplication
        'Dim rxsd As New SwitchPayIntegration.Models.Response
        'Dim hasMerchant As Boolean = False
        'Dim hasIDNumber As Boolean = False
        'Dim hasMobile As Boolean = False
        'Dim nofilters As Boolean = False
        'Dim results As Boolean = True
        'Dim amountvalid As Boolean = False
        'Dim bankvalid As Boolean = False
        'Dim mobilevalid As Boolean = True
        'Dim idvalid As Boolean = True
        'If (Not IsNumeric(IDNumber)) And (IDNumber <> String.Empty) Then
        '    idvalid = False
        'End If
        'If (Not IsNumeric(MobileNumber)) And (MobileNumber <> String.Empty) Then
        '    mobilevalid = False
        'End If
        'Try
        '    If (ApplicationRef = String.Empty) Or (IDNumber = String.Empty) Or (MobileNumber = String.Empty) Then
        '        nofilters = True
        '    End If
        '    If (FinanceAmount >= 500) And (FinanceAmount <= 230000) Then
        '        amountvalid = True
        '    End If
        '    If (Not IsNumeric(MobileNumber)) And (MobileNumber <> String.Empty) Then
        '        mobilevalid = False
        '    End If
        '    If (Not IsNumeric(IDNumber)) And (IDNumber <> String.Empty) Then
        '        idvalid = False
        '    End If
        '    Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        '    Try
        '        Dim bs = (From q In db.Banks Where q.ID = CLng(BankID) Select q).First()
        '        bankvalid = True
        '    Catch ex As Exception

        '    End Try
        '    If idvalid And bankvalid And mobilevalid And amountvalid And (Not nofilters) Then
        '        App.IDNumber = IDNumber
        '        App.MobileNumber = MobileNumber
        '        App.FinancialInstitutionID = BankID
        '        App.Reference = ApplicationRef
        '        App.InternalAuditTypeID = 1
        '        App.AuditTypeID = 1
        '        Dim applicationhistory As New Audit
        '        applicationhistory.Title = "New application created."
        '        applicationhistory.DateCreated = Now
        '        applicationhistory.Application = App
        '        applicationhistory.Entity = App.Entity
        '        db.Applications.InsertOnSubmit(App)
        '        db.SubmitChanges()
        '        Merchant = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
        '        hasMerchant = True
        '        Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
        '        hdt = rxsd.Tables("Header")
        '        Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
        '        hdr = hdt.NewHeaderRow()
        '        hdr("ApplicationID") = App.ID
        '        hdr("MerchantID") = MerchantID
        '        hdr("ApplicationRef") = ApplicationRef
        '        hdr("MerchantRef") = Merchant.Reference
        '        hdr("IDNumber") = IDNumber
        '        hdr("MobileNumber") = MobileNumber
        '        hdr("IsError") = False
        '        hdt.AddHeaderRow(hdr)
        '        Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
        '        fdt = rxsd.Tables("Fields")
        '        Dim fdr As SwitchPayIntegration.Models.Response.FieldsRow
        '        fdr = fdt.NewFieldsRow()
        '        fdr("Name") = "ApplicationID"
        '        fdr("Value") = App.ID
        '        fdt.AddFieldsRow(fdr)
        '        Dim fdr2 As SwitchPayIntegration.Models.Response.FieldsRow
        '        fdr2 = fdt.NewFieldsRow()
        '        fdr2("Name") = "Screen Message"
        '        fdr2("Value") = "Submitted"
        '        fdt.AddFieldsRow(fdr2)
        '        Dim fdr3 As SwitchPayIntegration.Models.Response.FieldsRow
        '        fdr3 = fdt.NewFieldsRow()
        '        fdr3("Name") = "SlipMessage"
        '        fdr3("Value") = "Application Submmitted"
        '        fdt.AddFieldsRow(fdr3)
        '        Dim fdr4 As SwitchPayIntegration.Models.Response.FieldsRow
        '        fdr4 = fdt.NewFieldsRow()
        '        fdr4("Name") = "TransactionType"
        '        fdr4("Value") = "Application"
        '        fdt.AddFieldsRow(fdr4)
        '        Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
        '        mdt = rxsd.Tables("Messages")
        '        Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
        '        mdr = mdt.NewMessagesRow()
        '        mdr("Code") = "0"
        '        mdr("Message") = "Application ID: " & App.ID & " successfully created"
        '        mdr("IsError") = False
        '        mdt.AddMessagesRow(mdr)
        '        If GenerateOTP Then
        '            Dim url = "https://www.xml2sms.gsm.co.za/send/?username=warpdev&password=Vodacom963&number=" & MobileNumber & "&message=Welcome to Switch Pay, your OTP is 11111"
        '            Dim client As New WebClient
        '            Dim Xml = client.DownloadString(url)
        '        End If


        '    Else
        '        Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
        '        hdt = rxsd.Tables("Header")
        '        Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
        '        hdr = hdt.NewHeaderRow()
        '        hdr("MerchantID") = MerchantID
        '        hdr("IsError") = True
        '        hdt.AddHeaderRow(hdr)
        '        Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
        '        mdt = rxsd.Tables("Messages")
        '        Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
        '        mdr2 = mdt.NewMessagesRow()
        '        mdr2("Code") = "5"
        '        mdr2("Message") = "Insufficient data provided."
        '        mdr2("IsError") = True
        '        mdt.AddMessagesRow(mdr2)
        '        If Not amountvalid Then
        '            Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
        '            mdr3 = mdt.NewMessagesRow()
        '            mdr3("Code") = "5"
        '            mdr3("Message") = "Amount Invalid"
        '            mdr3("IsError") = True
        '            mdt.AddMessagesRow(mdr3)

        '        End If
        '        If Not mobilevalid Then
        '            Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
        '            mdr3 = mdt.NewMessagesRow()
        '            mdr3("Code") = "5"
        '            mdr3("Message") = "Mobile Number Invalid"
        '            mdr3("IsError") = True
        '            mdt.AddMessagesRow(mdr3)

        '        End If
        '        If Not bankvalid Then
        '            Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
        '            mdr3 = mdt.NewMessagesRow()
        '            mdr3("Code") = "5"
        '            mdr3("Message") = "Bank Invalid"
        '            mdr3("IsError") = True
        '            mdt.AddMessagesRow(mdr3)

        '        End If
        '        If Not idvalid Then
        '            Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
        '            mdr3 = mdt.NewMessagesRow()
        '            mdr3("Code") = "5"
        '            mdr3("Message") = "ID Number Invalid"
        '            mdr3("IsError") = True
        '            mdt.AddMessagesRow(mdr3)

        '        End If
        '    End If
        'Catch ex As Exception
        '    Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
        '    hdt = rxsd.Tables("Header")
        '    'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        '    'Dim apps = (From q In db.Applications Where appl)
        '    Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
        '    hdr = hdt.NewHeaderRow()
        '    hdr("IsError") = True
        '    hdt.AddHeaderRow(hdr)
        '    Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
        '    fdt = rxsd.Tables("Fields")
        '    Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
        '    mdt = rxsd.Tables("Messages")
        '    If hasMerchant = False Then
        '        Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
        '        mdr3 = mdt.NewMessagesRow()
        '        mdr3("Code") = "4"
        '        mdr3("Message") = "No Valid Merchant Found."
        '        mdr3("IsError") = True
        '        mdt.AddMessagesRow(mdr3)

        '    End If
        '    Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
        '    mdr = mdt.NewMessagesRow()
        '    mdr("Code") = "1"
        '    mdr("Message") = "Exception: " & ex.Message
        '    mdr("IsError") = True
        '    mdt.AddMessagesRow(mdr)
        'End Try
        'Return rxsd
    End Function

    Function CreateLaybyApplication(MerchantID As Long, MerchantRef As String, TerminalID As String, ProductBandTermID As Long, ApplicationRef As String, FinanceAmount As Double, IDNumber As String, MobileNumber As String, BankID As Integer, GenerateOTP As Boolean, FirstName As String, Surname As String, GrossIncome As Double, NettIncome As Double, Term As Integer, Deposit As Double) As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.CreateLaybyApplication
        Dim rxsd As SwitchPayIntegration.Models.Response = AppCreation(MerchantID, String.Empty, ApplicationRef, TerminalID, FinanceAmount, IDNumber, MobileNumber, BankID, GenerateOTP, FirstName, Surname, GrossIncome, NettIncome, "LayBy", "Terminal")
        Dim fdt As SwitchPayIntegration.Models.Response.FieldsDataTable = rxsd.Tables("Fields")
        Dim fdr As SwitchPayIntegration.Models.Response.FieldsRow
        fdr = fdt.NewFieldsRow()
        fdr("Name") = "LeftColumn"
        fdr("Value") = "123|456|789|123|456"
        fdt.AddFieldsRow(fdr)
        Dim fdr2 As SwitchPayIntegration.Models.Response.FieldsRow
        fdr2 = fdt.NewFieldsRow()
        fdr2("Name") = "RightColumn"
        fdr2("Value") = "123|456|789|123|456"
        fdt.AddFieldsRow(fdr2)
        rxsd.AcceptChanges()
        Return rxsd
    End Function

    Function CreateMerchant(Title As String, ContactName As String, BankAccount As String, Reference As String, Phone As String, Email As String, ShortName As String, RegisteredName As String, RegNo As String, ParentMerchantID As Long) As Long Implements ISwitchPayAPI.CreateMerchant
        Dim rdb As New DBDataContext
        Dim m As New Merchant
        m.Title = Title
        m.DefaultBank = 9
        m.DefaultIDNo = String.Empty
        m.ID = CLng(RegisterMerchantSkelta(Reference, Title, Phone, "webconsumer.switchpay.co.za", Email, 11))
        If ParentMerchantID <> 0 Then
            m.ParentMerchantID = ParentMerchantID
        End If
        m.Reference = Reference
        m.ShortName = ShortName

        rdb.SubmitChanges()
        Return 1
    End Function
    '   <summary>
    '       This returns all lookup data
    '  </summary>


    Function CreateWorkflow(ApplicationID As Long, WorkflowName As String, QueueName As String, PriorityName As String) As String Implements ISwitchPayAPI.CreateWorkflow
        Try
            Dim dd As New DataDictionary(ApplicationID)
            Dim result = dd.WorkflowD.ExecuteWorkflow(WorkflowName, QueueName, PriorityName)
            Return result
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try

    End Function
    Public Sub DealRejected(ApplicationID As Long) Implements ISwitchPayAPI.DealRejected
        Try
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Dim a = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()

            Dim email2 As New MailMessage
            Dim SMTP As New SmtpClient("smtp.gmail.com")

            email2.From = New MailAddress("workflow@switchpay.co.za")
            SMTP.UseDefaultCredentials = False
            SMTP.Credentials = New System.Net.NetworkCredential("workflow@switchpay.co.za", "selfadrpcbiajyux") '<-- Password Here
            SMTP.EnableSsl = True
            email2.Subject = a.ID & " - Declined"
            email2.To.Add("support@acpas.co.za")
            email2.To.Add("diani@ammacom.com")
            email2.To.Add("Sacha.Craig@pmi.com")
            email2.To.Add("iqos@loanzie.co.za")

            email2.IsBodyHtml = True
            email2.Body = "Deal Declined<br />"
            Dim hist = (From q In db.vHistories Where q.ApplicationID = CLng(ApplicationID) Select q).ToArray()
            For Each h In hist
                email2.Body = email2.Body & h.AuditDate.ToString() & "<br />" & h.Details & "<br />"
            Next
            SMTP.Port = "587"
            SMTP.Send(email2)
            Dim dd As New DataDictionary(ApplicationID)
            Dim success As Boolean = dd.WorkflowD.RejectWF()
        Catch
        End Try

    End Sub

    Public Sub DeclineOffer(ApplicationID As Long) Implements ISwitchPayAPI.DeclineOffer
        Try
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Dim a = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
            Dim au As New Audit
            au.Name = "System"
            au.ApplicationID = a.ID
            au.AuditDate = Now

            au.Details = "Client Cancelled"
            au.AuditTypeID = 53
            db.Audits.InsertOnSubmit(au)
            db.SubmitChanges()
            a.AuditTypeID = 53
            db.SubmitChanges()
            Try
                Dim wfs As New List(Of String)
                wfs.Add("Additional Fields")
                ActionWFStep(ApplicationID, "Rejected", wfs)


            Catch ex As Exception
                'Message.Text &= ex.Message & ", "
            End Try
            Dim email2 As New MailMessage
            Dim SMTP As New SmtpClient("smtp.gmail.com")

            email2.From = New MailAddress("workflow@switchpay.co.za")
            SMTP.UseDefaultCredentials = False
            SMTP.Credentials = New System.Net.NetworkCredential("workflow@switchpay.co.za", "selfadrpcbiajyux") '<-- Password Here
            SMTP.EnableSsl = True
            email2.Subject = a.Reference & " - Cancelled"
            email2.To.Add("hendrik@acpas.co.za")
            email2.To.Add("jaco@acpas.co.za")
            email2.To.Add("support@acpas.co.za")
            email2.To.Add("diani@ammacom.com")
            email2.To.Add("Sacha.Craig@pmi.com")
            email2.To.Add("iqos@loanzie.co.za")

            email2.IsBodyHtml = True
            email2.Body = "Client Cancelled Deal<br />"
            Dim hist = (From q In db.vHistories Where q.ApplicationID = CLng(ApplicationID) Select q).ToArray()
            For Each h In hist
                email2.Body = email2.Body & h.AuditDate.ToString() & "<br />" & h.Details & "<br />"
            Next
            SMTP.Port = "587"
            SMTP.Send(email2)
            Dim dd As New DataDictionary(ApplicationID)
            Dim Success As Boolean = dd.WorkflowD.RejectWF()
        Catch
        End Try

    End Sub

    Public Sub DeleteBankDetails(ApplicationID) Implements ISwitchPayAPI.DeleteBankDetails
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim bs = (From q In db.ApplicationFieldValues Where ((q.FieldDefinitionEntityID = 250) Or (q.FieldDefinitionEntityID = 249) Or (q.FieldDefinitionEntityID = 248) Or (q.FieldDefinitionEntityID = 247)) And (q.ApplicationID = CLng(ApplicationID)) Select q).ToArray()
        For Each b In bs
            db.ApplicationFieldValues.DeleteOnSubmit(b)
        Next
        db.SubmitChanges()
    End Sub

    Public Sub FileUploaded(FileID As Long)
        Dim DataD As New DataDictionary(My.Settings.Environment, My.Settings.Repository)
        Dim p As New PaymentDictionary(DataD, "c:\templates\")
        p.FileUploaded(FileID)
    End Sub

    Public Sub GenerateCommInvoice(InvoiceID As Long) Implements ISwitchPayAPI.GenerateCommInvoice
        Dim DataD As New DataDictionary(My.Settings.Environment, My.Settings.Repository)
        Dim p As New PaymentDictionary(DataD, "c:\templates\")
        p.GenerateCommInvoice(InvoiceID)
    End Sub

    Public Sub GenerateDebitOrders(MerchantID As Long) Implements ISwitchPayAPI.GenerateDebitOrders
        Dim DataD As New DataDictionary(My.Settings.Environment, My.Settings.Repository)
        Dim p As New PaymentDictionary(DataD, "c:\templates\")
        p.GenerateDebitOrders(MerchantID)

    End Sub

    Public Sub GenerateInvoice(InvoiceID As Long) Implements ISwitchPayAPI.GenerateInvoice
        Dim DataD As New DataDictionary(My.Settings.Environment, My.Settings.Repository)
        Dim p As New PaymentDictionary(DataD, "c:\templates\")
        p.GenerateInvoice(InvoiceID)


    End Sub

    Public Sub GeneratePaymentRequest(ApplicationID As Long) Implements ISwitchPayAPI.GeneratePaymentRequest
        Dim DataD As New DataDictionary(ApplicationID)
        ' Dim p As New PaymentDictionary(DataD, "c:\templates\")
        DataD.PaymentD.GeneratePaymentRequest(ApplicationID)
    End Sub

    Public Function GenerateSubsFile(FileID As Long) As String Implements ISwitchPayAPI.GenerateSubsFile
        Dim DataD As New DataDictionary(My.Settings.Environment, My.Settings.Repository)
        Dim p As New PaymentDictionary(DataD, "c:\templates\")
        p.GenerateSubsFile(FileID)
#Disable Warning BC42105 ' Function 'GenerateSubsFile' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.
    End Function
#Enable Warning BC42105 ' Function 'GenerateSubsFile' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.

    Function GetAdditionalFields(ApplicationID As Long, Responsestr As String) As String Implements ISwitchPayAPI.GetAdditionalFields
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
        Dim apps = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) Select q).ToArray()
        'Dim ser As JObject = JObject.Parse(Response)
        Dim str As String = String.Empty
        For Each a In apps
            str &= "<AdditionalInfo>
                       <FieldTitle>" & (From q In rdb.FieldDefinitionEntities Where q.ID = a.FieldDefinitionEntityID).First.Title & "</FieldTitle>
                       <FieldValue>" & a.Title & "</FieldValue>
                   </AdditionalInfo>
"
        Next
        str = Responsestr.Replace("<AdditionalInfo></AdditionalInfo>", str)


        Return str
    End Function

    Public Function GetApplicationFieldValue(ApplicationID As Long, FieldDefinitionEntityID As Long) As String Implements ISwitchPayAPI.GetApplicationFieldValue
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)

        Try
            Dim afv = (From q In db.ApplicationFieldValues Where (q.ApplicationID = CLng(ApplicationID)) And (q.FieldDefinitionEntityID = CLng(FieldDefinitionEntityID)) Select q.Title).First()
            Return afv
        Catch
            Return "Not Supplied"
        End Try
    End Function

    Public Function GetApplicationFieldValueCode(ApplicationID As Long, FieldDefinitionEntityID As Long) As String Implements ISwitchPayAPI.GetApplicationFieldValueCode
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
        Try
            Dim afv = (From q In db.ApplicationFieldValues Where (q.ApplicationID = CLng(ApplicationID)) And (q.FieldDefinitionEntityID = CLng(FieldDefinitionEntityID)) Select q).First()
            Dim fo = (From q In rdb.FieldOptions Where (q.FieldDefinitionID = CLng((From w In rdb.FieldDefinitionEntities Where w.ID = afv.FieldDefinitionEntityID).First.FieldDefinitionID)) And (q.Title.ToUpper() = CStr(afv.Title).ToUpper()) Select q.Code).First()
            Return fo
        Catch
            Return "Not Supplied"
        End Try
    End Function


    Public Function GetApplications(MerchantID As Long, TerminalID As String, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.GetApplications
        Dim rxsd As New SwitchPayIntegration.Models.Response
        Dim hasMerchant As Boolean = False
        Dim hasIDNumber As Boolean = False
        Dim hasMobile As Boolean = False
        Dim nofilters As Boolean = False
        Dim results As Boolean = True
        Dim amountvalid As Boolean = False
        Dim bankvalid As Boolean = False
        Dim mobilevalid As Boolean = True
        Dim idvalid As Boolean = True
        Dim alreadycancelled As Boolean = False
        If (Not IsNumeric(IDNumber)) And (IDNumber <> String.Empty) Then
            idvalid = False
        End If
        If (Not IsNumeric(MobileNumber)) And (MobileNumber <> String.Empty) Then
            mobilevalid = False
        End If
        Try
            If (ApplicationID = 0) And (ApplicationRef = String.Empty) And (IDNumber = String.Empty) And (MobileNumber = String.Empty) Then
                nofilters = True
            End If
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)

            Dim Apps As New List(Of Application)
            If ApplicationID <> 0 Then
                Apps = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).ToList()
            ElseIf ApplicationRef <> String.Empty Then
                Apps = (From q In db.Applications Where q.Reference = CStr(ApplicationRef) Select q).ToList()
            ElseIf IDNumber <> String.Empty Then
                Apps = (From q In db.Applications Where q.IDNumber = CStr(IDNumber) Select q).ToList()
            Else
                'Apps = (From q In db.Applications Where q.entityi = CStr(ApplicationRef) Select q).ToArray
            End If
            If Apps.Count = 0 Then
                results = False
            End If
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            'Dim apps = (From q In db.Applications Where appl)
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            For Each Appl In Apps
                hdr = hdt.NewHeaderRow()
                hdr("ApplicationID") = App.ID
                hdr("MerchantID") = MerchantID
                hdr("ApplicationRef") = App.Reference
                Merchant = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
                hdr("MerchantRef") = Merchant.Reference
                hasMerchant = True
                hdr("IDNumber") = App.IDNumber
                hdr("MobileNumber") = App.MobileNumber
                hdr("IsError") = False
                hdt.AddHeaderRow(hdr)
            Next
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If Not results Then
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "2"
                mdr2("Message") = "No Applications Found."
                mdr2("IsError") = True
                mdt.AddMessagesRow(mdr2)
            End If
            If nofilters Then
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "3"
                mdr2("Message") = "No Filters Provided."
                mdr2("IsError") = True
                mdt.AddMessagesRow(mdr2)
            End If
            If Not idvalid Then
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "5"
                mdr2("Message") = "No Valid ID Number Provided."
                mdr2("IsError") = True
                mdt.AddMessagesRow(mdr2)
            End If
            If Not mobilevalid Then
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "3"
                mdr2("Message") = "No Valid Mobile Number Provided."
                mdr2("IsError") = True
                mdt.AddMessagesRow(mdr2)
            End If
        Catch ex As Exception
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            'Dim apps = (From q In db.Applications Where appl)
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            hdr = hdt.NewHeaderRow()
            hdr("IsError") = True
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If hasMerchant = False Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "4"
                mdr3("Message") = "No Valid Merchant Found."
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "1"
            mdr("Message") = "Exception: " & ex.Message
            mdr("IsError") = True
            mdt.AddMessagesRow(mdr)
        End Try
        Return rxsd
    End Function

    Function GetApplicationStatus(MerchantID As Long, TerminalID As String, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.GetApplicationStatus
        Dim rxsd As New SwitchPayIntegration.Models.Response
        Dim hasMerchant As Boolean = False
        Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
        Try
            Dim nofilters As Boolean = False
            Dim results As Boolean = True
            If (ApplicationID = 0) And (ApplicationRef = String.Empty) And (IDNumber = String.Empty) And (MobileNumber = String.Empty) Then
                nofilters = True
            End If
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Merchant = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
            hasMerchant = True
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            Try
                If ApplicationID <> 0 Then
                    App = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
                ElseIf IDNumber <> String.Empty Then
                    App = (From q In db.Applications Where q.IDNumber = CStr(IDNumber) Select q Order By q.DateCreated Descending).First()
                ElseIf ApplicationRef <> String.Empty Then
                    App = (From q In db.Applications Where q.Reference = CStr(ApplicationRef) Select q Order By q.DateCreated Descending).First()
                Else
                    'Apps = (From q In db.Applications Where q.entityi = CStr(ApplicationRef) Select q).ToArray
                    results = False
                End If
            Catch
                results = False
            End Try
            hdr = hdt.NewHeaderRow()
            If results Then
                hdr("ApplicationID") = App.ID
                hdr("MerchantID") = MerchantID
                hdr("ApplicationRef") = App.Reference
                Try
                    Merchant = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
                    hasMerchant = True
                    hdr("MerchantRef") = Merchant.Reference
                Catch
                End Try
                hdr("IDNumber") = App.IDNumber
                hdr("MobileNumber") = App.MobileNumber
                hdr("IsError") = False
                Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
                fdt = rxsd.Tables("Fields")
                Dim fdr111 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr111 = fdt.NewFieldsRow()
                fdr111("Name") = "Collectable"
                If CanCollect(App.ID) Then
                    fdr111("Value") = True
                Else
                    fdr111("Value") = False
                End If
                fdt.AddFieldsRow(fdr111)
                Dim fdr222 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr222 = fdt.NewFieldsRow()
                fdr222("Name") = "MaxAmount"
                fdr222("Value") = 230000
                fdt.AddFieldsRow(fdr222)
                Dim fdr333 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr333 = fdt.NewFieldsRow()
                fdr333("Name") = "MinAmount"
                fdr333("Value") = 500
                fdt.AddFieldsRow(fdr333)
                Dim fdr444 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr444 = fdt.NewFieldsRow()
                fdr444("Name") = "TestAmount"
                fdr444("Value") = 500
                fdt.AddFieldsRow(fdr444)
                Dim fdr As SwitchPayIntegration.Models.Response.FieldsRow
                fdr = fdt.NewFieldsRow()
                fdr("Name") = "StatusID"
                fdr("Value") = App.AuditTypeID
                fdt.AddFieldsRow(fdr)
                Dim fdr2 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr2 = fdt.NewFieldsRow()
                fdr2("Name") = "Status"
                fdr2("Value") = (From q In rdb.AuditTypes Where q.ID = App.AuditTypeID).First().Title ' .au.Title
                fdt.AddFieldsRow(fdr2)
                Dim fdr5 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr5 = fdt.NewFieldsRow()
                fdr5("Name") = "Screen Message"
                fdr5("Value") = (From q In rdb.AuditTypes Where q.ID = App.AuditTypeID).First().Title
                fdt.AddFieldsRow(fdr5)
                Dim fdr3 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr3 = fdt.NewFieldsRow()
                fdr3("Name") = "SlipMessage"
                fdr3("Value") = "Application Status: " & (From q In rdb.AuditTypes Where q.ID = App.AuditTypeID).First().Title
                fdt.AddFieldsRow(fdr3)
                Dim fdr4 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr4 = fdt.NewFieldsRow()
                fdr4("Name") = "TransactionType"
                fdr4("Value") = "Status Check"
                fdt.AddFieldsRow(fdr4)
                Dim fdr9 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr9 = fdt.NewFieldsRow()
                fdr9("Name") = "MerchantName"
                fdr9("Value") = Merchant.Title
                fdt.AddFieldsRow(fdr9)
                Dim fdr11 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr11 = fdt.NewFieldsRow()
                fdr11("Name") = "MerchantActive"
                Try
                    Dim mes = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
                    fdr11("Value") = mes.IsActive
                    fdt.AddFieldsRow(fdr11)
                Catch
                    fdr11("Value") = False
                    fdt.AddFieldsRow(fdr11)
                End Try
                Try
                    Dim fdr6 As SwitchPayIntegration.Models.Response.FieldsRow
                    fdr6 = fdt.NewFieldsRow()
                    fdr6("Name") = "AvailableBalance"
                    fdr6("Value") = String.Format("{0:0.00}", App.OfferAmount)
                    fdt.AddFieldsRow(fdr6)
                Catch
                End Try
                Dim fdr9999 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr9999 = fdt.NewFieldsRow()
                fdr9999("Name") = "CollectionDate"
                Try
                    Dim vApp = (From q In db.vApplications Where q.ID = ApplicationID Select q).FirstOrDefault()
                    fdr9999("Value") = vApp.DateCollected

                Catch ex As Exception
                    fdr9999("Value") = DBNull.Value

                End Try
                fdt.AddFieldsRow(fdr9999)
            Else
                hdr("IsError") = True
            End If
            hdt.AddHeaderRow(hdr)
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If nofilters Then
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "3"
                mdr2("Message") = "No Application Found."
                mdr2("IsError") = True
                mdt.AddMessagesRow(mdr2)
            ElseIf Not results Then
                Dim mdr4 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr4 = mdt.NewMessagesRow()
                mdr4("Code") = "2"
                mdr4("Message") = "No Applications Found."
                mdr4("IsError") = True
                mdt.AddMessagesRow(mdr4)
            Else
                Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
                mdr = mdt.NewMessagesRow()
                mdr("Code") = "0"
                mdr("Message") = "Application ID: " & App.ID & " successfully checked status"
                mdr("IsError") = False
                mdt.AddMessagesRow(mdr)
            End If
        Catch ex As Exception
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            'Dim apps = (From q In db.Applications Where appl)
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            hdr = hdt.NewHeaderRow()
            hdr("IsError") = True
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If hasMerchant = False Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "4"
                mdr3("Message") = "No Valid Merchant Found."
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "1"
            mdr("Message") = "Exception: " & ex.Message
            mdr("IsError") = True
            mdt.AddMessagesRow(mdr)
        End Try
        Return rxsd
    End Function

    Function GetApplicationStatusFailed(MerchantID As Long, TerminalID As String, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As SwitchPayIntegration.Models.Response
        Dim rxsd As New SwitchPayIntegration.Models.Response
        Dim hasMerchant As Boolean = False
        Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
        Try
            Dim nofilters As Boolean = False
            Dim results As Boolean = True
            If (ApplicationID = 0) And (ApplicationRef = String.Empty) And (IDNumber = String.Empty) And (MobileNumber = String.Empty) Then
                nofilters = True
            End If
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Merchant = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
            hasMerchant = True
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            Try
                If ApplicationID <> 0 Then
                    App = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
                ElseIf IDNumber <> String.Empty Then
                    App = (From q In db.Applications Where q.IDNumber = CStr(IDNumber) Select q Order By q.DateCreated Descending).First()
                ElseIf ApplicationRef <> String.Empty Then
                    App = (From q In db.Applications Where q.Reference = CStr(ApplicationRef) Select q Order By q.DateCreated Descending).First()
                Else
                    'Apps = (From q In db.Applications Where q.entityi = CStr(ApplicationRef) Select q).ToArray
                    results = False
                End If
            Catch
                results = False
            End Try
            hdr = hdt.NewHeaderRow()
            If results Then
                hdr("ApplicationID") = App.ID
                hdr("MerchantID") = MerchantID
                hdr("ApplicationRef") = App.Reference
                Try
                    Merchant = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
                    hasMerchant = True
                    hdr("MerchantRef") = Merchant.Reference
                Catch
                End Try
                hdr("IDNumber") = App.IDNumber
                hdr("MobileNumber") = App.MobileNumber
                hdr("IsError") = False
                Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
                fdt = rxsd.Tables("Fields")
                Dim fdr111 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr111 = fdt.NewFieldsRow()
                fdr111("Name") = "Collectable"
                If CanCollect(App.ID) Then
                    fdr111("Value") = True
                Else
                    fdr111("Value") = False
                End If
                fdt.AddFieldsRow(fdr111)
                Dim fdr222 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr222 = fdt.NewFieldsRow()
                fdr222("Name") = "MaxAmount"
                fdr222("Value") = 230000
                fdt.AddFieldsRow(fdr222)
                Dim fdr333 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr333 = fdt.NewFieldsRow()
                fdr333("Name") = "MinAmount"
                fdr333("Value") = 500
                fdt.AddFieldsRow(fdr333)
                Dim fdr444 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr444 = fdt.NewFieldsRow()
                fdr444("Name") = "TestAmount"
                fdr444("Value") = 500
                fdt.AddFieldsRow(fdr444)
                Dim fdr As SwitchPayIntegration.Models.Response.FieldsRow
                fdr = fdt.NewFieldsRow()
                fdr("Name") = "StatusID"
                fdr("Value") = App.AuditTypeID
                fdt.AddFieldsRow(fdr)
                Dim fdr2 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr2 = fdt.NewFieldsRow()
                fdr2("Name") = "Status"
                fdr2("Value") = (From q In rdb.AuditTypes Where q.ID = App.AuditTypeID).First().Title ' .au.Title
                fdt.AddFieldsRow(fdr2)
                Dim fdr5 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr5 = fdt.NewFieldsRow()
                fdr5("Name") = "Screen Message"
                fdr5("Value") = (From q In rdb.AuditTypes Where q.ID = App.AuditTypeID).First().Title
                fdt.AddFieldsRow(fdr5)
                Dim fdr3 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr3 = fdt.NewFieldsRow()
                fdr3("Name") = "SlipMessage"
                fdr3("Value") = "Application Status: " & (From q In rdb.AuditTypes Where q.ID = App.AuditTypeID).First().Title
                fdt.AddFieldsRow(fdr3)
                Dim fdr4 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr4 = fdt.NewFieldsRow()
                fdr4("Name") = "TransactionType"
                fdr4("Value") = "Status Check"
                fdt.AddFieldsRow(fdr4)
                Dim fdr9 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr9 = fdt.NewFieldsRow()
                fdr9("Name") = "MerchantName"
                fdr9("Value") = Merchant.Title
                fdt.AddFieldsRow(fdr9)
                Dim fdr11 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr11 = fdt.NewFieldsRow()
                fdr11("Name") = "MerchantActive"
                Try
                    Dim mes = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
                    fdr11("Value") = mes.IsActive
                    fdt.AddFieldsRow(fdr11)
                Catch
                    fdr11("Value") = False
                    fdt.AddFieldsRow(fdr11)
                End Try
                Try
                    Dim fdr6 As SwitchPayIntegration.Models.Response.FieldsRow
                    fdr6 = fdt.NewFieldsRow()
                    fdr6("Name") = "AvailableBalance"
                    fdr6("Value") = String.Format("{0:0.00}", App.OfferAmount)
                    fdt.AddFieldsRow(fdr6)
                Catch
                End Try
                Dim fdr9999 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr9999 = fdt.NewFieldsRow()
                fdr9999("Name") = "CollectionDate"
                Try
                    Dim vApp = (From q In db.vApplications Where q.ID = ApplicationID Select q).FirstOrDefault()
                    fdr9999("Value") = vApp.DateCollected

                Catch ex As Exception
                    fdr9999("Value") = DBNull.Value

                End Try
                fdt.AddFieldsRow(fdr9999)
            Else
                hdr("IsError") = True
            End If
            hdt.AddHeaderRow(hdr)
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If nofilters Then
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "3"
                mdr2("Message") = "No Filters Provided."
                mdr2("IsError") = True
                mdt.AddMessagesRow(mdr2)
            ElseIf Not results Then
                Dim mdr4 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr4 = mdt.NewMessagesRow()
                mdr4("Code") = "2"
                mdr4("Message") = "No Applications Found."
                mdr4("IsError") = True
                mdt.AddMessagesRow(mdr4)
            Else
                Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
                mdr = mdt.NewMessagesRow()
                mdr("Code") = "0"
                mdr("Message") = "Application ID: " & App.ID & " successfully checked status"
                mdr("IsError") = False
                mdt.AddMessagesRow(mdr)
            End If
        Catch ex As Exception
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            'Dim apps = (From q In db.Applications Where appl)
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            hdr = hdt.NewHeaderRow()
            hdr("IsError") = True
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If hasMerchant = False Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "4"
                mdr3("Message") = "No Valid Merchant Found."
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "1"
            mdr("Message") = "Exception: " & ex.Message
            mdr("IsError") = True
            mdt.AddMessagesRow(mdr)
        End Try
        Return rxsd
    End Function

    Function GetApplicationStatusShort(ByVal ApplicationID As Long) As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.GetApplicationStatusShort
        Dim rxsd As New SwitchPayIntegration.Models.Response
        Dim hasMerchant As Boolean = False
        Try
            Dim nofilters As Boolean = False
            Dim results As Boolean = True
            If ApplicationID = 0 Then
                nofilters = True
            End If
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            Try
                App = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()

            Catch
                results = False
            End Try
            hdr = hdt.NewHeaderRow()
            If results Then
                Merchant = (From q In rdb.Merchants Where q.ID = CLng(App.MerchantID) Select q).First()
                hasMerchant = True
                hdr("ApplicationID") = App.ID
                hdr("MerchantID") = App.MerchantID
                hdr("ApplicationRef") = App.Reference
                Try
                    Merchant = (From q In rdb.Merchants Where q.ID = CLng(App.MerchantID) Select q).First()
                    hasMerchant = True
                    hdr("MerchantRef") = Merchant.Reference
                Catch
                End Try
                hdr("IDNumber") = App.IDNumber
                hdr("MobileNumber") = App.MobileNumber
                hdr("IsError") = False
                Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
                fdt = rxsd.Tables("Fields")
                Dim fdr111 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr111 = fdt.NewFieldsRow()
                fdr111("Name") = "Collectable"
                If CanCollect(App.ID) Then
                    fdr111("Value") = True
                Else
                    fdr111("Value") = False
                End If
                fdt.AddFieldsRow(fdr111)
                Dim fdr222 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr222 = fdt.NewFieldsRow()
                fdr222("Name") = "MaxAmount"
                fdr222("Value") = 230000
                fdt.AddFieldsRow(fdr222)
                Dim fdr333 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr333 = fdt.NewFieldsRow()
                fdr333("Name") = "MinAmount"
                fdr333("Value") = 500
                fdt.AddFieldsRow(fdr333)
                Dim fdr444 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr444 = fdt.NewFieldsRow()
                fdr444("Name") = "TestAmount"
                fdr444("Value") = 500
                fdt.AddFieldsRow(fdr444)
                Dim fdr As SwitchPayIntegration.Models.Response.FieldsRow
                fdr = fdt.NewFieldsRow()
                fdr("Name") = "StatusID"
                fdr("Value") = App.AuditTypeID
                fdt.AddFieldsRow(fdr)
                Dim fdr2 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr2 = fdt.NewFieldsRow()
                fdr2("Name") = "Status"
                fdr2("Value") = (From q In rdb.AuditTypes Where q.ID = App.AuditTypeID Select q.Title).First
                fdt.AddFieldsRow(fdr2)
                Dim fdr5 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr5 = fdt.NewFieldsRow()
                fdr5("Name") = "Screen Message"
                fdr5("Value") = (From q In rdb.AuditTypes Where q.ID = App.AuditTypeID Select q.Title).First
                fdt.AddFieldsRow(fdr5)
                Dim fdr3 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr3 = fdt.NewFieldsRow()
                fdr3("Name") = "SlipMessage"
                fdr3("Value") = "Application Status: " & (From q In rdb.AuditTypes Where q.ID = App.AuditTypeID Select q.Title).First
                fdt.AddFieldsRow(fdr3)
                Dim fdr4 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr4 = fdt.NewFieldsRow()
                fdr4("Name") = "TransactionType"
                fdr4("Value") = "Status Check"
                fdt.AddFieldsRow(fdr4)
                Dim fdr9 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr9 = fdt.NewFieldsRow()
                fdr9("Name") = "MerchantName"
                fdr9("Value") = Merchant.Title
                fdt.AddFieldsRow(fdr9)
                Dim fdr11 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr11 = fdt.NewFieldsRow()
                fdr11("Name") = "MerchantActive"
                Try
                    Dim mes = (From q In rdb.Merchants Where q.ID = CLng(App.MerchantID) Select q).First()
                    fdr11("Value") = mes.IsActive
                    fdt.AddFieldsRow(fdr11)
                Catch
                    fdr11("Value") = False
                    fdt.AddFieldsRow(fdr11)
                End Try
                Try
                    Dim fdr6 As SwitchPayIntegration.Models.Response.FieldsRow
                    fdr6 = fdt.NewFieldsRow()
                    fdr6("Name") = "AvailableBalance"
                    fdr6("Value") = String.Format("{0:c}", Math.Round(CDbl(App.OfferAmount), 2).ToString())
                    fdt.AddFieldsRow(fdr6)
                Catch
                End Try
            Else
                hdr("IsError") = True
            End If
            hdt.AddHeaderRow(hdr)
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If nofilters Then
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "3"
                mdr2("Message") = "No Application Found."
                mdr2("IsError") = True
                mdt.AddMessagesRow(mdr2)
            ElseIf Not results Then
                Dim mdr4 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr4 = mdt.NewMessagesRow()
                mdr4("Code") = "2"
                mdr4("Message") = "No Applications Found."
                mdr4("IsError") = True
                mdt.AddMessagesRow(mdr4)
            Else
                Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
                mdr = mdt.NewMessagesRow()
                mdr("Code") = "0"
                mdr("Message") = "Application ID: " & App.ID & " successfully checked status"
                mdr("IsError") = False
                mdt.AddMessagesRow(mdr)
            End If
        Catch ex As Exception
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            'Dim apps = (From q In db.Applications Where appl)
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            hdr = hdt.NewHeaderRow()
            hdr("IsError") = True
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If hasMerchant = False Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "4"
                mdr3("Message") = "No Valid Merchant Found."
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "1"
            mdr("Message") = "Exception: " & ex.Message
            mdr("IsError") = True
            mdt.AddMessagesRow(mdr)
        End Try
        Return rxsd
    End Function

    Public Function GetBanksByMerchantID(MerchantID As Long, MerchantRef As String, TerminalID As String, Purpose As Integer) As DataTable Implements ISwitchPayAPI.GetBanksByMerchantID
        Dim db As New RulesDBDataContext(My.Settings.RulesDB)
        Dim bs = (From q In db.FinancialInstitutions Where q.IsDeleted = False Select q.ID, q.Title, q.Phone, q.Email, q.IsCreditProvider, q.IsOnBankList, q.IDForTerminal, q.IDForBankList).ToList()
        Dim dt As New DataTable("Banks")
        dt.Columns.Add("ID", GetType(Long))
        dt.Columns.Add("Title", GetType(String))
        dt.Columns.Add("Phone", GetType(String))
        dt.Columns.Add("Email", GetType(String))
        dt.Columns.Add("IsCreditProvider", GetType(Boolean))
        dt.Columns.Add("IsOnBankList", GetType(Boolean))
        dt.Columns.Add("IDForTerminal", GetType(Long))
        dt.Columns.Add("IDForBankList", GetType(Long))
        dt.AcceptChanges()
        For Each b In bs
            dt.Rows.Add(New Object() {b.ID, b.Title, b.Phone, b.Email, b.IsCreditProvider, b.IsOnBankList, b.IDForTerminal, b.IDForBankList})
        Next
        dt.AcceptChanges()
        Return dt
    End Function

    Public Function GetLastTransaction(MerchantID As Long, MerchantRef As String, TerminalID As String, Purpose As Integer) As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.GetLastTransaction
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim Rdb As New RulesDBDataContext(My.Settings.RulesDB)
        Dim rxsd As New SwitchPayIntegration.Models.Response
        Dim hasMerchant As Boolean = False
        Try
            If MerchantRef = String.Empty Then
                Merchant = (From q In Rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
                hasMerchant = True
            Else
                Merchant = (From q In Rdb.Merchants Where q.Reference = CStr(MerchantRef) Select q).First()
                hasMerchant = True
            End If
            If hasMerchant Then
                rxsd = GetApplicationStatus(Merchant.ID, TerminalID, 0, String.Empty, String.Empty, String.Empty)

                Return rxsd
            Else
                Merchant = (From q In Rdb.Merchants Where q.ID = 1 Select q).First()
                Dim term = (From q In Rdb.MerchantTerminals Where q.MerchantID = 1 Select q).First()

                rxsd = GetApplicationStatus(Merchant.ID, TerminalID, 0, String.Empty, String.Empty, String.Empty)
                Return rxsd
            End If
        Catch
            Merchant = (From q In Rdb.Merchants Where q.ID = 1 Select q).First()
            Dim term = (From q In Rdb.MerchantTerminals Where q.MerchantID = 1 Select q).First()

            rxsd = GetApplicationStatus(Merchant.ID, TerminalID, 0, String.Empty, String.Empty, String.Empty)
            Return rxsd
        End Try
    End Function

    Public Function GetLayByDetails(MerchantID As Long, MerchantRef As String, TerminalID As String, Amount As Double, ProductBandTermID As Long, ProductID As Long) As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.GetLayByDetails
        Try
            Dim rxsd As New SwitchPayIntegration.Models.Response
            Dim mdt As SwitchPayIntegration.Models.Response.MessagesDataTable = rxsd.Tables("Messages")
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "101"
            mdr("Message") = "Amount = R" & String.Format("{0:0.00}", Amount)
            mdr("IsError") = False
            mdt.AddMessagesRow(mdr)
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "102"
            Dim Deposit As Double = Amount * 0.2
            mdr("Message") = "Deposit = R" & String.Format("{0:0.00}", Deposit)
            mdr("IsError") = False
            mdt.AddMessagesRow(mdr)
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "103"
            mdr("Message") = "Activation Fee = R" & String.Format("{0:0.00}", 70)
            mdr("IsError") = False
            mdt.AddMessagesRow(mdr)
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "104"
            Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
            Dim terms = (From q In rdb.ProductBandTerms Where q.ID = ProductBandTermID Select q.Title).FirstOrDefault()
            mdr("Message") = "Payment Term = " & terms & " months"
            mdr("IsError") = False
            mdt.AddMessagesRow(mdr)
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "105"
            Dim Balance As Double = Amount - Deposit
            mdr("Message") = "Total After Deposit = R" & String.Format("{0:0.00}", Balance)
            mdr("IsError") = False
            mdt.AddMessagesRow(mdr)
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "106"
            Dim Payment As Double = Balance / terms
            mdr("Message") = "Monthly Payment = R" & String.Format("{0:0.00}", Payment)
            mdr("IsError") = False
            mdt.AddMessagesRow(mdr)
            rxsd.AcceptChanges()
            Return rxsd
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Function GetLookUps(MerchantID As Long, MerchantRef As String, TerminalID As String, Purpose As Integer) As SwitchPayIntegration.Models.LookUps Implements ISwitchPayAPI.GetLookUps
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim Rdb As New RulesDBDataContext(My.Settings.RulesDB)
        Dim bxsd As New SwitchPayIntegration.Models.LookUps
        Dim bdt As New SwitchPayIntegration.Models.LookUps.BankDataTable
        bdt = bxsd.Tables("Bank")
        Dim bs = (From q In Rdb.FinancialInstitutions Where q.IsDeleted = False Select q.ID, q.Title, q.Phone, q.Email, q.IsCreditProvider, q.IsOnBankList, q.IDForTerminal, q.IDForBankList).ToList()
        For Each b In bs

            Dim br As SwitchPayIntegration.Models.LookUps.BankRow
            br = bdt.NewBankRow()
            br("ID") = b.ID
            br("Title") = b.Title
            br("Phone") = b.Phone
            br("Email") = b.Email
            br("IsCreditProvider") = b.IsCreditProvider
            br("IsOnBankList") = b.IsOnBankList
            Try
                br("IDForTerminal") = b.IDForTerminal
            Catch
                br("IDForTerminal") = DBNull.Value
            End Try
            Try
                br("IDForBankList") = b.IDForBankList
            Catch
                br("IDForBankList") = DBNull.Value
            End Try
            bdt.AddBankRow(br)
        Next
        Dim sdt As New SwitchPayIntegration.Models.LookUps.StatusDataTable
        sdt = bxsd.Tables("Status")
        Dim ss = (From q In Rdb.AuditTypes Select q).ToArray()
        For Each s As AuditType In ss

            Dim sr As SwitchPayIntegration.Models.LookUps.StatusRow
            sr = sdt.NewStatusRow()
            sr("ID") = s.ID
            sr("Title") = s.Title
            Try
                sr("Description") = s.Description
            Catch
                sr("Description") = DBNull.Value
            End Try
            sdt.AddStatusRow(sr)
        Next
        Dim idt As New SwitchPayIntegration.Models.LookUps.IndustryDataTable
        idt = bxsd.Tables("Industry")
        Dim iss = (From q In Rdb.AuditTypes Select q).ToArray()
        For Each i In iss

            Dim ir As SwitchPayIntegration.Models.LookUps.IndustryRow
            ir = idt.NewIndustryRow()
            ir("ID") = i.ID
            ir("Title") = i.Title
            idt.AddIndustryRow(ir)
        Next
        bxsd.AcceptChanges()
        Return bxsd
    End Function

    Public Function GetMerchantLogo(MerchantID As Long) As Byte() Implements ISwitchPayAPI.GetMerchantLogo
        'Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
        'Dim Application As ApplicationObject = New ApplicationObject(My.Settings.WFApp)
        'Dim av As New AttachmentSecuredValue
        'Dim data As Byte()
        'Dim m = (From q In rdb.vMerchants Where q.nvarchar53 = CStr(MerchantID) Select q).First()
        'av.ResolveCurrentValue(m.image2.ToString())


        ''Dim l As New Skelta.Repository.List.ListDefinition(Application, New Guid("76471462-BC17-4715-A465-093C609888F9"))
        ''Dim li As New Skelta.Repository.List.ListItem(l, m.Id)
        ''Dim lf As New Skelta.Repository.List.Field
        'Return data
#Disable Warning BC42105 ' Function 'GetMerchantLogo' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.
    End Function
#Enable Warning BC42105 ' Function 'GetMerchantLogo' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.

    Public Function GetMerchantName(MerchantID As Long) As String Implements ISwitchPayAPI.GetMerchantName
        Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
        Dim m = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
        Return m.ShortName
    End Function

    Function GetMerchantStatus(Optional ByVal MerchantID As Long = 0, Optional ByVal MerchantRef As String = "", Optional ByVal TerminalID As String = "") As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.GetMerchantStatus
        Dim rxsd As New SwitchPayIntegration.Models.Response
        Dim hasMerchant As Boolean = False
        Try
            Dim dd As New DataDictionary
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            hdr = hdt.NewHeaderRow()
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Dim cdb As New DBDataContext(My.Settings.ChildDB)
            If MerchantRef = String.Empty Then
                dd.LoadMerchantByID(MerchantID, TerminalID)
                hasMerchant = True
            Else
                dd.LoadMerchantByReference(MerchantRef, TerminalID)
                hasMerchant = True
            End If
            hdr("MerchantID") = dd.MID
            hdr("MerchantRef") = dd.AppMerchant.Reference
            Try
                If dd.AppMerchant.IsActive Then
                    hdr("MerchantActive") = "True"
                Else
                    hdr("MerchantActive") = "False"
                End If
            Catch
                hdr("MerchantActive") = "False"
            End Try
            Try
                If dd.AppMerchantTerminal.IsActive Then
                    hdr("TerminalActive") = "True"
                Else
                    hdr("TerminalActive") = "False"
                End If
            Catch
                hdr("TerminalActive") = "False"
            End Try
            hdr("MerchantRef") = dd.AppMerchant.Reference
            hdr("IsError") = False
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim fdr As SwitchPayIntegration.Models.Response.FieldsRow
            fdr = fdt.NewFieldsRow()
            fdr("Name") = "Terminals"
            Try
                fdr("Value") = dd.MerchantD.MerchantTerminals.Count
            Catch
                fdr("Value") = 0
            End Try
            fdt.AddFieldsRow(fdr)

            fdr = fdt.NewFieldsRow()
            fdr("Name") = "StatusID"
            fdr("Value") = 2
            fdt.AddFieldsRow(fdr)

            Dim fdr2 As SwitchPayIntegration.Models.Response.FieldsRow
            fdr2 = fdt.NewFieldsRow()
            fdr2("Name") = "Status"
#Disable Warning BC42024 ' Unused local variable: 'mes'.
            Dim mes As Merchant
#Enable Warning BC42024 ' Unused local variable: 'mes'.
            Try
                Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
                fdr2("Value") = dd.AppMerchant.IsActive And dd.AppMerchantTerminal.IsActive
                fdt.AddFieldsRow(fdr2)
            Catch
                fdr2("Value") = False
                fdt.AddFieldsRow(fdr2)
            End Try
            Dim fdr57 As SwitchPayIntegration.Models.Response.FieldsRow
            fdr57 = fdt.NewFieldsRow()
            fdr57("Name") = "ProductIDs"
            fdr57("Value") = "1|2"
            fdt.AddFieldsRow(fdr57)
            fdr57 = fdt.NewFieldsRow()
            fdr57("Name") = "ProductNames"
            fdr57("Value") = "Finance|LayBy"
            fdt.AddFieldsRow(fdr57)
            fdr57 = fdt.NewFieldsRow()
            fdr57("Name") = "Range1"
            fdr57("Value") = "0|1500"
            fdt.AddFieldsRow(fdr57)
            fdr57 = fdt.NewFieldsRow()
            fdr57("Name") = "Term1"
            fdr57("Value") = "3|6"
            fdt.AddFieldsRow(fdr57)
            fdr57 = fdt.NewFieldsRow()
            fdr57("Name") = "Range2"
            fdr57("Value") = "1501|5000"
            fdt.AddFieldsRow(fdr57)
            fdr57 = fdt.NewFieldsRow()
            fdr57("Name") = "Term2"
            fdr57("Value") = "6|9|12"
            fdt.AddFieldsRow(fdr57)

            Try
                Dim fdr7 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr7 = fdt.NewFieldsRow()
                fdr7("Name") = "MerchantName"
                fdr7("Value") = dd.AppMerchant.Title
                fdt.AddFieldsRow(fdr7)
                Dim fdr77 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr77 = fdt.NewFieldsRow()
                fdr77("Name") = "MerchantDisplayName"
                fdr77("Value") = dd.AppMerchant.ShortName
                fdt.AddFieldsRow(fdr77)
            Catch
                Dim fdr7 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr7 = fdt.NewFieldsRow()
                fdr7("Name") = "MerchantName"
                fdr7("Value") = dd.AppMerchant.Title
                fdt.AddFieldsRow(fdr7)
                Dim fdr77 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr77 = fdt.NewFieldsRow()
                fdr77("Name") = "MerchantDisplayName"
                fdr77("Value") = dd.AppMerchant.ShortName
                fdt.AddFieldsRow(fdr77)
            End Try
            Dim fdr8 As SwitchPayIntegration.Models.Response.FieldsRow
            fdr8 = fdt.NewFieldsRow()
            fdr8("Name") = "Address"
            Try
                fdr8("Value") = dd.AppMerchantDetails.PhysicalAddress & "|" & dd.AppMerchantDetails.PhysicalSuburb & " - " & dd.AppMerchantDetails.PhysicalCity & "|" & dd.AppMerchantDetails.PhysicalProvince & "|South Africa"
            Catch
                fdr8("Value") = " | | | "
            End Try
            fdt.AddFieldsRow(fdr8)
            Dim fdr9 As SwitchPayIntegration.Models.Response.FieldsRow
            fdr9 = fdt.NewFieldsRow()
            fdr9("Name") = "EndPoint"

            fdr9("Value") = dd.EnvironmentD.DestinationEnvironment.URLPrefix & dd.ExternalWebAPIURL.DNS & dd.ExternalWebAPIURL.PostFix

            fdt.AddFieldsRow(fdr9)
            Dim fdr10 As SwitchPayIntegration.Models.Response.FieldsRow
            fdr10 = fdt.NewFieldsRow()
            fdr10("Name") = "MinAmount"
            fdr10("Value") = 500
            fdt.AddFieldsRow(fdr10)
            Dim fdr11 As SwitchPayIntegration.Models.Response.FieldsRow
            fdr11 = fdt.NewFieldsRow()
            fdr11("Name") = "MaxAmount"
            fdr11("Value") = 300000
            fdt.AddFieldsRow(fdr11)
            Dim fdr12 As SwitchPayIntegration.Models.Response.FieldsRow
            fdr12 = fdt.NewFieldsRow()
            fdr12("Name") = "DefaultAmount"
            fdr12("Value") = dd.AppMerchant.DefaultAmount
            fdt.AddFieldsRow(fdr12)
            Dim fdr13 As SwitchPayIntegration.Models.Response.FieldsRow
            fdr13 = fdt.NewFieldsRow()
            fdr13("Name") = "DefaultIDNo"
            fdr13("Value") = dd.AppMerchant.DefaultIDNo  '"1111111111111|2222222222222|3333333333333"
            fdt.AddFieldsRow(fdr13)
            Dim fdr15 As SwitchPayIntegration.Models.Response.FieldsRow
            fdr15 = fdt.NewFieldsRow()
            fdr15("Name") = "DefaultMobile"
            fdr15("Value") = dd.AppMerchant.DefaultMobile
            fdt.AddFieldsRow(fdr15)
            Dim fdr14 As SwitchPayIntegration.Models.Response.FieldsRow
            fdr14 = fdt.NewFieldsRow()
            fdr14("Name") = "DefaultBank"
            fdr14("Value") = dd.AppMerchant.DefaultBank
            fdt.AddFieldsRow(fdr14)
            Dim fdr16 As SwitchPayIntegration.Models.Response.FieldsRow
            fdr16 = fdt.NewFieldsRow()
            fdr16("Name") = "DefaultOTP"
            fdr16("Value") = dd.AppMerchant.DefaultOTP
            fdt.AddFieldsRow(fdr16)
            Dim fdr99 As SwitchPayIntegration.Models.Response.FieldsRow
            fdr99 = fdt.NewFieldsRow()
            fdr99("Name") = "Environment"

            fdr99("Value") = dd.DestinationEnvironment.Title

            fdt.AddFieldsRow(fdr99)
            Dim fdr98 As SwitchPayIntegration.Models.Response.FieldsRow
            fdr98 = fdt.NewFieldsRow()
            fdr98("Name") = "ConnectIP"
            If dd.DestinationEnvironment.Title = "UAT" Then

                fdr98("Value") = "10.255.0.16"
                'ElseIf dd.DestinationEnvironment.Title = "Staging" Then
                '    fdr98("Value") = "stagingwebapi.switchpay.co.za"
            Else
                fdr98("Value") = "10.255.1.16"
            End If

            fdt.AddFieldsRow(fdr98)
            Dim fdr97 As SwitchPayIntegration.Models.Response.FieldsRow
            fdr97 = fdt.NewFieldsRow()
            fdr97("Name") = "Port"
            Select Case dd.DestinationEnvironment.Title
                Case "Production"
                    fdr97("Value") = "36001"
                Case "UAT"
                    fdr97("Value") = "35001"
                Case Else
                    fdr97("Value") = "36001"
            End Select

            fdt.AddFieldsRow(fdr97)
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "0"
            mdr("Message") = "Merchant ID: " & dd.MID & " status successfully checked"
            mdr("IsError") = False
            mdt.AddMessagesRow(mdr)
            dd.MerchantD.AddMerchantHistory(5, "Status Checked: Merchant ID: " & MerchantID & " Merchant Ref: " & MerchantRef & " Terminal ID: " & TerminalID)

        Catch ex As Exception
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            'Dim apps = (From q In db.Applications Where appl)
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            hdr = hdt.NewHeaderRow()
            hdr("IsError") = True
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If hasMerchant = False Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "4"
                mdr3("Message") = "No Valid Merchant Found."
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "1"
            mdr("Message") = "Exception: " & ex.Message
            mdr("IsError") = True
            mdt.AddMessagesRow(mdr)
        End Try
        Return rxsd
    End Function

    Function GetTerms(MerchantID As Long, MerchantRef As String, TerminalID As String, Amount As Double, ProductID As Long) As SwitchPayIntegration.Models.Products Implements ISwitchPayAPI.GetTerms
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim Rdb As New RulesDBDataContext(My.Settings.RulesDB)
        Dim bxsd1 As New SwitchPayIntegration.Models.Products
        Dim bdt As New SwitchPayIntegration.Models.Products.ProductBandTermsDataTable
        bdt = bxsd1.Tables("ProductBandTerm")
        '  Dim bs = (From q In Rdb.ProductBandTerms Where q.pr  q.IsDeleted = False Select q.ID, q.Title).ToList
        For i = 1 To 3
            Dim br As SwitchPayIntegration.Models.Products.ProductBandTermsRow
            br = bdt.NewProductBandTermsRow()
            br("ID") = i
            br("Title") = (i * 3).ToString()
            br("ProductBandID") = 1
            bdt.AddProductBandTermsRow(br)
        Next
        bxsd1.AcceptChanges()
        Return bxsd1
    End Function

    Public Function isActivity(ApplicationID As Long, DisplayName As String) As String Implements ISwitchPayAPI.isActivity
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)

#Disable Warning BC42024 ' Unused local variable: 'a'.
        Dim a As Application
#Enable Warning BC42024 ' Unused local variable: 'a'.
        Dim result As Boolean = False
        Try
            Dim dd As New DataDictionary(ApplicationID)

            result = dd.WorkflowD.IsWFActivity("Authorise", DisplayName, result)
        Catch
        End Try
        Return result.ToString()
    End Function

    Public Sub LoadMerchantDetailsData(MerchantID As Long)
        Try
            Merchant = New Merchant
        Catch

        End Try
    End Sub

    Public Function LoanzieAcceptOffer(ApplicationID As Long) As Decimal Implements ISwitchPayAPI.LoanzieAcceptOffer
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim a = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
        Dim dd As New DataDictionary

        Try
            Dim l As New AP.externalintegrationSoapClient
            Dim PayDate As Date
            PayDate = New Date(Now.AddMonths(1).Year, Now.AddMonths(1).Month, Math.Min(CInt(GetApplicationFieldValue(a.ID, 245)), Date.DaysInMonth(Now.AddMonths(1).Year, Now.AddMonths(1).Month)))

            Dim cust = l.Do_Customer_Exists(dd.ACPASUser.Title, dd.ACPASUser.Password, dd.ACPASUser.Code, a.IDNumber)
            If a.Reference.Contains(cust) Then
            Else
                Dim d = l.Insert_New_Agreement_With_Correlation(dd.ACPASUser.Title, dd.ACPASUser.Password, dd.ACPASUser.Code, cust, a.OfferAmount, PayDate, 11, 82, 84, 4, 0, 43, a.ID.ToString() & "_" & a.MerchantID.ToString(), "", "")
                'CDate(Now.AddDays((Now.Day - 1) * -1).AddMonths(1)), 11, 82, 84, 4, 0, 43)
                db = New DBDataContext(My.Settings.SwitchPayDB)
                a = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
                a.Reference = cust & "," & d
                db.SubmitChanges()
                Threading.Thread.Sleep(1000)
                Dim f = l.Pop_Customer_Agreement(dd.ACPASUser.Title, dd.ACPASUser.Password, dd.ACPASUser.Code, cust, d)
                a.OfferInstallment = f.Tables(0).Rows(0)("INSTALMENT")
                a.Reference = cust & "," & f.Tables(0).Rows(0)("AGREEMENTNO")
                db.SubmitChanges()
            End If
            Dim au As New Audit
            au.Name = "System"
            au.ApplicationID = a.ID
            au.AuditDate = Now
            au.Details = "Deal Accepted for: " & String.Format("{0:0.00}", a.OfferAmount) & "Term: " & a.OfferTerm & "Installment: " & String.Format("{0:0.00}", a.OfferInstallment)
            a.AuditTypeID = 10
            au.AuditTypeID = 10
            db.SubmitChanges()
            db.Audits.InsertOnSubmit(au)
            db.SubmitChanges()
            a = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
            Return a.OfferAmount
        Catch ex As Exception
            a = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
            Dim au As New Audit
            au.Name = "System"
            au.ApplicationID = a.ID
            au.AuditDate = Now
            au.Details = "Error Occurred: " & ex.Message
            a.AuditTypeID = 19
            au.AuditTypeID = 19
            db.SubmitChanges()
            db.Audits.InsertOnSubmit(au)
            db.SubmitChanges()
            ' SendSMS(a.ID, "Fin App " & a.ID & " has been declined by Loanzie. Queries 0105945332.")
            Throw ex
        End Try
    End Function


    Public Function LoanzieGetAgreement(ApplicationID As Long) As DataSet Implements ISwitchPayAPI.LoanzieGetAgreement
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim a = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
        Dim dd As New DataDictionary

        Try
            Dim l As New AP.externalintegrationSoapClient

            Dim results = Split(a.Reference, ",")
            Dim F = l.Pop_Customer_Agreement(dd.ACPASUser.Title, dd.ACPASUser.Password, dd.ACPASUser.Code, results(0), results(1))

            Return F
        Catch ex As Exception
            Dim au As New Audit
            au.Name = "System"
            au.ApplicationID = a.ID
            au.AuditDate = Now
            au.Details = "Deal Declined: " & ex.Message
            a.AuditTypeID = 11
            au.AuditTypeID = 11
            db.SubmitChanges()
            db.Audits.InsertOnSubmit(au)
            db.SubmitChanges()
            SendSMS(a.ID, "Fin App " & a.ID & " has been declined by Loanzie. Queries 0105945332.")
            Return Nothing
        End Try
    End Function

    Public Function grossapproved(ID As Long) As Double Implements ISwitchPayAPI.grossapproved
        Dim db As New DBDataContext
        Dim app2 = (From q In db.Applications Where q.ID = CLng(ID) Select q).First()
        Dim salary = CDbl(GetApplicationFieldValue(ID, 253))
        Dim approved As Boolean = False

        If salary >= 4041.5 And app2.FinanceAmount <= 969.96 Then
            approved = True
            app2.OfferAmount = Math.Round(CDbl((app2.FinanceAmount / 12) * 11), 2)
        End If
        If salary >= 5416.5 And app2.FinanceAmount <= 1299.96 Then
            approved = True
            app2.OfferAmount = Math.Round(CDbl((app2.FinanceAmount / 12) * 11), 2)
        End If
        If salary >= 5708.5 And App.FinanceAmount <= 1370.04 Then
            approved = True
            app2.OfferAmount = Math.Round(CDbl((app2.FinanceAmount / 12) * 11), 2)
        End If
        If salary >= 9791.5 And app2.FinanceAmount <= 2349.96 Then
            approved = True
            app2.OfferAmount = Math.Round(CDbl((app2.FinanceAmount / 12) * 11), 2)
        End If
        If approved Then
            db.SubmitChanges()
        Else
            app2.OfferAmount = 0
            db.SubmitChanges()
        End If
        Return app2.OfferAmount
    End Function

    Public Function CapitecPrevet(ApplicationID As Long) As String Implements ISwitchPayAPI.CapitecPrevet
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim a = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
        Dim dd As New DataDictionary

        Try
            Dim l As New AP.externalintegrationSoapClient
            Dim countagree As Integer = 0
            Try
                Dim cust = l.Do_Customer_Exists(dd.ACPASUser.Title, dd.ACPASUser.Password, dd.ACPASUser.Code, a.IDNumber)

                If cust <> 0 Then
                    Dim appcount = (From q In db.vApplications Where (q.ID <> a.ID) And (q.IDNumber = a.IDNumber) And q.Approved Select q).Count()
                    countagree = appcount
                End If
            Catch ex As Exception

            End Try
            Dim approved As Boolean = False
            a.OfferAmount = Math.Round(CDbl((a.FinanceAmount / 12) * 11), 2)
            Dim Year = Convert.ToInt32(a.IDNumber.Substring(0, 2))
            Dim Month = Convert.ToInt32(a.IDNumber.Substring(2, 2))
            Dim Day = Convert.ToInt32(a.IDNumber.Substring(4, 2))

            Dim dob As Date = New Date(IIf(Year > 25, 1900 + Year, 2000 + Year), Month, Day)

            Dim xx = l.XDS_PreVetting_Non_Client(dd.ACPASUser.Title, dd.ACPASUser.Password, dd.ACPASUser.Code, a.FirstName, a.Surname, a.IDNumber)
            If DateDiff(DateInterval.Year, dob, CDate(Now)) < 18 Then
                Dim au As New Audit
                au.Name = "System"
                au.ApplicationID = a.ID
                au.AuditDate = Now
                au.Details = "Loanzie Score: " & xx.EnquiryScore.ToString() & ", Decision: " & xx.EnquiryDecision.ToUpper() & ", Deal Declined - UnderAge"
                a.AuditTypeID = 11
                a.OfferAmount = 0
                au.AuditTypeID = 11
                db.SubmitChanges()
                db.Audits.InsertOnSubmit(au)
                db.SubmitChanges()
                SendSMS(a.ID, "Fin App " & a.ID & " has been declined by Loanzie, under 18. Queries 0105945332.")
                DealRejected(a.ID)
                Return "0"

            End If

            Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
            Try
                approved = True
                If xx.EnquiryDecision.ToUpper() = "FAIL" Then
                    Dim au As New Audit
                    au.Name = "System"
                    au.ApplicationID = a.ID
                    au.AuditDate = Now
                    au.Details = "Loanzie Score: " & xx.EnquiryScore.ToString() & ", Decision: " & xx.EnquiryDecision.ToUpper() & ", Deal Declined - " & xx.EnquiryExculsionReason.ToString()
                    a.AuditTypeID = 11
                    a.OfferAmount = 0
                    au.AuditTypeID = 11
                    db.SubmitChanges()
                    db.Audits.InsertOnSubmit(au)
                    db.SubmitChanges()
                    SendSMS(a.ID, "Fin App " & a.ID & " has been declined by Loanzie. Queries 0105945332.")
                    DealRejected(a.ID)
                    Return "0"
                End If
            Catch
                Dim au As New Audit
                au.Name = "System"
                au.ApplicationID = a.ID
                au.AuditDate = Now
                au.Details = "Loanzie Score: " & xx.EnquiryScore.ToString() & ", Decision: " & xx.EnquiryDecision.ToUpper() & ", Deal Declined - Error calling ACPAS"
                a.AuditTypeID = 11
                a.OfferAmount = 0
                au.AuditTypeID = 11
                db.SubmitChanges()
                db.Audits.InsertOnSubmit(au)
                db.SubmitChanges()
                SendSMS(a.ID, "Fin App " & a.ID & " has been declined by Loanzie. Queries 0105945332.")
                DealRejected(a.ID)
                Return "0"
            End Try
            a.OfferTerm = 11
            a.OfferInstallment = Math.Round(CDbl(a.OfferAmount) / 11, 2)
            a.AuditTypeID = 4
            db.SubmitChanges()
            Dim au2 As New Audit
            au2.Name = "System"
            au2.ApplicationID = a.ID
            au2.AuditDate = Now
            au2.Details = "Loanzie Score: " & xx.EnquiryScore.ToString() & ", Decision: " & xx.EnquiryDecision.ToUpper() & "Provisional Offer sent for: " & String.Format("{0:0.00}", a.OfferAmount) & "Term: " & a.OfferTerm & "Installment: " & String.Format("{0:0.00}", a.OfferInstallment)
            au2.AuditTypeID = 4
            db.Audits.InsertOnSubmit(au2)
            db.SubmitChanges()
            Return a.OfferAmount.ToString()
        Catch ex As Exception
            Dim au As New Audit
            au.Name = "System"
            au.ApplicationID = a.ID
            au.AuditDate = Now
            au.Details = "Deal Declined: " & ex.Message
            a.AuditTypeID = 11
            a.OfferAmount = 0
            au.AuditTypeID = 11
            db.SubmitChanges()
            db.Audits.InsertOnSubmit(au)
            db.SubmitChanges()
            SendSMS(a.ID, "Fin App " & a.ID & " has been declined by Loanzie. Queries 0105945332.")
            DealRejected(a.ID)
            Return "0"
        End Try
    End Function

    Public Function LoanziePrevet(ApplicationID As Long) As String Implements ISwitchPayAPI.LoanziePrevet
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim a = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
        Dim dd As New DataDictionary

        Try
            Dim l As New AP.externalintegrationSoapClient
            Dim countagree As Integer = 0
            Try
                Dim cust = l.Do_Customer_Exists(dd.ACPASUser.Title, dd.ACPASUser.Password, dd.ACPASUser.Code, a.IDNumber)

                If cust <> 0 Then
                    Dim appcount = (From q In db.vApplications Where (q.ID <> a.ID) And (q.IDNumber = a.IDNumber) And q.Approved Select q).Count()
                    countagree = appcount
                End If
            Catch ex As Exception

            End Try
            'If countagree >= 3 Then
            '    Dim au As New Audit
            '    au.Name = "System"
            '    au.ApplicationID = a.ID
            '    au.AuditDate = Now
            '    au.Details = "Deal Declined - Too Many Agreements"
            '    a.AuditTypeID = 11
            '    a.OfferAmount = 0
            '    au.AuditTypeID = 11
            '    db.SubmitChanges()
            '    db.Audits.InsertOnSubmit(au)
            '    db.SubmitChanges()
            '    SendSMS(a.ID, "Fin App " & a.ID & " has been declined by Loanzie, too many active agreements. Queries 0105945332.")
            '    DealRejected(a.ID)
            '    Return "0"
            'End If
            Dim approved As Boolean = False
            a.OfferAmount = Math.Round(CDbl((a.FinanceAmount / 12) * 11), 2)
            Dim Year = Convert.ToInt32(a.IDNumber.Substring(0, 2))
            Dim Month = Convert.ToInt32(a.IDNumber.Substring(2, 2))
            Dim Day = Convert.ToInt32(a.IDNumber.Substring(4, 2))

            Dim dob As Date = New Date(IIf(Year > 25, 1900 + Year, 2000 + Year), Month, Day)

            Dim xx = l.XDS_PreVetting_Non_Client(dd.ACPASUser.Title, dd.ACPASUser.Password, dd.ACPASUser.Code, a.FirstName, a.Surname, a.IDNumber)
            If DateDiff(DateInterval.Year, dob, CDate(Now)) < 18 Then
                Dim au As New Audit
                au.Name = "System"
                au.ApplicationID = a.ID
                au.AuditDate = Now
                au.Details = "Loanzie Score: " & xx.EnquiryScore.ToString() & ", Decision: " & xx.EnquiryDecision.ToUpper() & ", Deal Declined - UnderAge"
                a.AuditTypeID = 11
                a.OfferAmount = 0
                au.AuditTypeID = 11
                db.SubmitChanges()
                db.Audits.InsertOnSubmit(au)
                db.SubmitChanges()
                SendSMS(a.ID, "Fin App " & a.ID & " has been declined by Loanzie, under 18. Queries 0105945332.")
                DealRejected(a.ID)
                Return "0"

            End If

            Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
            Try
                approved = True
                If xx.EnquiryDecision.ToUpper() = "FAIL" Then
                    Dim au As New Audit
                    au.Name = "System"
                    au.ApplicationID = a.ID
                    au.AuditDate = Now
                    au.Details = "Loanzie Score: " & xx.EnquiryScore.ToString() & ", Decision: " & xx.EnquiryDecision.ToUpper() & ", Deal Declined - " & xx.EnquiryExculsionReason.ToString()
                    a.AuditTypeID = 11
                    a.OfferAmount = 0
                    au.AuditTypeID = 11
                    db.SubmitChanges()
                    db.Audits.InsertOnSubmit(au)
                    db.SubmitChanges()
                    SendSMS(a.ID, "Fin App " & a.ID & " has been declined by Loanzie. Queries 0105945332.")
                    DealRejected(a.ID)
                    Return "0"
                End If
            Catch
                Dim au As New Audit
                au.Name = "System"
                au.ApplicationID = a.ID
                au.AuditDate = Now
                au.Details = "Loanzie Score: " & xx.EnquiryScore.ToString() & ", Decision: " & xx.EnquiryDecision.ToUpper() & ", Deal Declined - Error calling ACPAS"
                a.AuditTypeID = 11
                a.OfferAmount = 0
                au.AuditTypeID = 11
                db.SubmitChanges()
                db.Audits.InsertOnSubmit(au)
                db.SubmitChanges()
                SendSMS(a.ID, "Fin App " & a.ID & " has been declined by Loanzie. Queries 0105945332.")
                DealRejected(a.ID)
                Return "0"
            End Try
            a.OfferTerm = 11
            a.OfferInstallment = Math.Round(CDbl(a.OfferAmount) / 11, 2)
            a.AuditTypeID = 4
            db.SubmitChanges()
            Dim au2 As New Audit
            au2.Name = "System"
            au2.ApplicationID = a.ID
            au2.AuditDate = Now
            au2.Details = "Loanzie Score: " & xx.EnquiryScore.ToString() & ", Decision: " & xx.EnquiryDecision.ToUpper() & "Provisional Offer sent for: " & String.Format("{0:0.00}", a.OfferAmount) & "Term: " & a.OfferTerm & "Installment: " & String.Format("{0:0.00}", a.OfferInstallment)
            au2.AuditTypeID = 4
            db.Audits.InsertOnSubmit(au2)
            db.SubmitChanges()
            Return a.OfferAmount.ToString()
        Catch ex As Exception
            Dim au As New Audit
            au.Name = "System"
            au.ApplicationID = a.ID
            au.AuditDate = Now
            au.Details = "Deal Declined: " & ex.Message
            a.AuditTypeID = 11
            a.OfferAmount = 0
            au.AuditTypeID = 11
            db.SubmitChanges()
            db.Audits.InsertOnSubmit(au)
            db.SubmitChanges()
            SendSMS(a.ID, "Fin App " & a.ID & " has been declined by Loanzie. Queries 0105945332.")
            DealRejected(a.ID)
            Return "0"
        End Try
    End Function

    Function LoanzieQuickCheck(ApplicationID As Long) As String Implements ISwitchPayAPI.LoanzieQuickCheck
        '		Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        '		Dim app = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
        '		Dim appid As String
        '		If Not (app.LoanzieGuid Is Nothing) Then
        '			If app.LoanzieGuid.Length > 0 Then
        '				appid = app.LoanzieGuid
        '			Else
        '				appid = Guid.NewGuid().ToString()
        '				app.LoanzieGuid = appid
        '				db.SubmitChanges()

        '			End If
        '		Else
        '			appid = Guid.NewGuid().ToString()
        '			app.LoanzieGuid = appid
        '			db.SubmitChanges()
        '		End If
        '		Dim f As New Loanzie
        '		Dim str As String = "{
        '  ""appNr"": """ & appid & """,
        '  ""consumer"": {
        '	""name"": ""ALWYN"",
        '	""surname"": ""BADENHORST"",
        '	""idNo"": ""7101015244089"",
        '	""idType"": ""R"",
        '  },
        '  ""accountType"": ""LZE_S"",
        '  ""dealAmount"": " & app.FinanceAmount.ToString().Replace(",", ".") & ",
        '  ""grossIncome"": " & app.GrossIncome.ToString().Replace(",", ".") & ",
        '  ""nettIncome"": " & app.NettIncome.ToString().Replace(",", ".") & ",
        '  ""otherIncome"": 0,
        '  ""expenses"": [

        '  ]
        '}"

        '		Dim token As String = f.QuickCheck(str)

        '		Dim json As String = token
        '		Dim ser As JObject = JObject.Parse(json)
        '		Dim utoken As String = ser("appId")
        '		Return utoken
#Disable Warning BC42105 ' Function 'LoanzieQuickCheck' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.
    End Function
#Enable Warning BC42105 ' Function 'LoanzieQuickCheck' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.

    Function LoanzieQuickCheckOffer(ApplicationID As Long, Response As String) As String Implements ISwitchPayAPI.LoanzieQuickCheckOffer
        '		Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        '		Dim app = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()


        '		Dim str2 As String = "{
        '  ""appId"": """ & Response & """,
        '  ""consumer"": {
        '	""name"": ""ALWYN"",
        '	""surname"": ""BADENHORST"",
        '	""idNo"": ""7101015244089"",
        '  ""idType"": ""R"",
        '  },
        '  ""accountType"": ""LZE_S"",
        ' ""dealAmount"": " & app.FinanceAmount.ToString().Replace(",", ".") & ",
        '  ""grossIncome"": " & app.GrossIncome.ToString().Replace(",", ".") & ",
        '  ""nettIncome"": " & app.NettIncome.ToString().Replace(",", ".") & ",
        '  ""otherIncome"": 0,
        '  ""expenses"": [

        '  ]
        '}"



        '		Dim f As New Loanzie
        '		Dim ustring As String = f.QuickCheckOffer(str2)
        '		Dim json As String = ustring
        '		Dim ser2 As JObject = JObject.Parse(json)

        '		Dim jt As JToken = ser2("offers").Children().Last()
        '		ser2 = JObject.Parse(jt.ToString())
        '		Dim utoken2 As String = jt.ToString()
        '		app.OfferAmount = ser2("loanAmount")
        '		app.OfferTerm = ser2("term")
        '		app.OfferInstallment = ser2("instalment")
        '		db.SubmitChanges()
        '		utoken2 = ApplicationID & "[2"
        '		Return utoken2
#Disable Warning BC42105 ' Function 'LoanzieQuickCheckOffer' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.
    End Function
#Enable Warning BC42105 ' Function 'LoanzieQuickCheckOffer' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.

    Public Sub LoanzieSaveData(ApplicationID As Long) Implements ISwitchPayAPI.LoanzieSaveData
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim a = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
        Dim dd As New DataDictionary

        Try
            Dim l As New AP.externalintegrationSoapClient

            Dim cust = l.Do_Customer_Exists(dd.ACPASUser.Title, dd.ACPASUser.Password, dd.ACPASUser.Code, a.IDNumber)
            If cust = 0 Then


                Dim Year = Convert.ToInt32(a.IDNumber.Substring(0, 2))
                Dim Month = Convert.ToInt32(a.IDNumber.Substring(2, 2))
                Dim Day = Convert.ToInt32(a.IDNumber.Substring(4, 2))

                Dim dob As Date = New Date(IIf(Year > 25, 1900 + Year, 2000 + Year), Month, Day)
                cust = l.Insert_New_Customer(dd.ACPASUser.Title, dd.ACPASUser.Password, dd.ACPASUser.Code, a.FirstName, a.Surname, a.FirstName.Substring(0, 1), a.IDNumber, CDate(dob), 7, 2, 5, "1 Lime Street", "Bryanston", "Bryanston", "Gauteng", 1, "2196", "1 Lime Street", "Bryanston", "Bryanston", "Gauteng", 1, "2196", GetApplicationFieldValue(a.ID, 243), a.MobileNumber, a.MobileNumber, String.Empty, True, 27, 10, a.ID, a.FinanceAmount)
                ' Dim af = l.Insert_Customer_Affordability(dd.ACPASUser.Title, dd.ACPASUser.Password, dd.ACPASUser.Code, cust, GetApplicationFieldValue(a.ID, 3), 0, 0, 0, 0, CDbl(GetApplicationFieldValue(a.ID, 3) - GetApplicationFieldValue(a.ID, 3), GetApplicationFieldValue(a.ID, 3), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2000, 2000, 8000)
            End If
            Dim c = l.Insert_Customer_Bank(dd.ACPASUser.Title, dd.ACPASUser.Password, dd.ACPASUser.Code, cust, GetApplicationFieldValueCode(a.ID, 248), GetApplicationFieldValue(a.ID, 247), GetApplicationFieldValue(a.ID, 249), GetApplicationFieldValueCode(a.ID, 250), 12, 2024, GetApplicationFieldValueCode(a.ID, 248))
            'Dim y3 = l.Get_EmployerList(dd.ACPASUser.Title, dd.ACPASUser.Password, dd.ACPASUser.Code, 17)

            Dim b As Long
            'Dim empname As String = GetApplicationFieldValue(a.ID, 244)
            ''Dim found As Boolean = False
            ''For Each dr In y3.Tables(0).Rows
            ''    If dr(1).ToString.ToUpper = empname.ToUpper Then
            ''        b = dr(0)
            ''        found = True
            ''        Exit For
            ''    End If

            ''Next
            ''If Not found Then
            b = l.Insert_New_Employer(dd.ACPASUser.Title, dd.ACPASUser.Password, dd.ACPASUser.Code, GetApplicationFieldValue(a.ID, 244), "1 Lime Street", "Bryanston", "Bryanston", "Gauteng", 1, "2196", "1 Lime Street", "Bryanston", "Bryanston", "Gauteng", 1, "2196", "Bryan", GetApplicationFieldValue(a.ID, 246), GetApplicationFieldValue(a.ID, 246), "bryan@switchpay.co.za", GetApplicationFieldValue(a.ID, 245), 21)
            'End If

            l.Insert_Customer_Employer(dd.ACPASUser.Title, dd.ACPASUser.Password, dd.ACPASUser.Code, cust, b, "HR", "Nothing", "123")
        Catch ex As Exception
            Dim au As New Audit
            au.Name = "System"
            au.ApplicationID = a.ID
            au.AuditDate = Now
            au.Details = "Error Saving Loanzie Data: " & ex.Message
            au.AuditTypeID = 19
            db.SubmitChanges()
            db.Audits.InsertOnSubmit(au)
            db.SubmitChanges()
            Throw ex
        End Try
    End Sub

    Function NIUSSDAuthorise(ApplicationID As Long, Response As String) As String Implements ISwitchPayAPI.NIUSSDAuthorise
        Dim result As String
        result = SendNIUSSD(ApplicationID, Response)
        Return result
    End Function

    Public Sub NTUOffer(ApplicationID As Long) Implements ISwitchPayAPI.NTUOffer
        Try
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Dim a = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
            Dim au As New Audit
            au.Name = "System"
            au.ApplicationID = a.ID
            au.AuditDate = Now

            au.Details = "Offer Expired"
            au.AuditTypeID = 39
            db.Audits.InsertOnSubmit(au)
            db.SubmitChanges()
            a.AuditTypeID = 39
            db.SubmitChanges()
            Dim email2 As New MailMessage
            Dim SMTP As New SmtpClient("smtp.gmail.com")

            email2.From = New MailAddress("workflow@switchpay.co.za")
            SMTP.UseDefaultCredentials = False
            SMTP.Credentials = New System.Net.NetworkCredential("workflow@switchpay.co.za", "selfadrpcbiajyux") '<-- Password Here
            SMTP.EnableSsl = True
            email2.Subject = a.Reference & " - Cancelled Expired"
            email2.To.Add("hendrik@acpas.co.za")
            email2.To.Add("jaco@acpas.co.za")
            email2.To.Add("support@acpas.co.za")
            email2.To.Add("diani@ammacom.com")
            email2.To.Add("Sacha.Craig@pmi.com")
            email2.To.Add("iqos@loanzie.co.za")

            email2.IsBodyHtml = True
            email2.Body = "Client Did Not Take Up Deal<br />"
            Dim hist = (From q In db.vHistories Where q.ApplicationID = CLng(ApplicationID) Select q).ToArray()
            For Each h In hist
                email2.Body = email2.Body & h.AuditDate.ToString() & "<br />" & h.Details & "<br />"
            Next
            SMTP.Port = "587"
            SMTP.Send(email2)
            Dim dd As New DataDictionary(ApplicationID)
            dd.WorkflowD.RejectWF()
        Catch
        End Try

    End Sub

    Public Sub PaidCommFile(FileID As Long, Items As String) Implements ISwitchPayAPI.PaidCommFile
        Dim DataD As New DataDictionary(My.Settings.Environment, My.Settings.Repository)
        Dim p As New PaymentDictionary(DataD, "c:\templates\")
        p.PaidCommFile(FileID, Items)
    End Sub


    Public Sub PaidFile(FileID As Long, Items As String) Implements ISwitchPayAPI.PaidFile
        Dim DataD As New DataDictionary(My.Settings.Environment, My.Settings.Repository)
        Dim p As New PaymentDictionary(DataD, "c:\templates\")
        p.PaidFile(FileID, Items)
    End Sub

    Function Podium(Message As String) As Boolean Implements ISwitchPayAPI.Podium

        Return True
    End Function


    Public Function PreparePaymentRun(ApplicationID As Long) Implements ISwitchPayAPI.PreparePaymentRun
        Dim DataD As New DataDictionary(ApplicationID)
        Dim p As New PaymentDictionary(DataD, "c:\templates\")
        p.PreparePaymentRun(ApplicationID)
#Disable Warning BC42105 ' Function 'PreparePaymentRun' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.
    End Function
#Enable Warning BC42105 ' Function 'PreparePaymentRun' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.

    Public Sub PreparePaymentsFile(FileID As Long) Implements ISwitchPayAPI.PreparePaymentsFile
        Dim DataD As New DataDictionary(My.Settings.Environment, My.Settings.Repository)
        Dim p As New PaymentDictionary(DataD, "c:\templates\")
        p.PreparePaymentsFile(FileID)
    End Sub

    Public Function ReceiveDeliveryReceipt(AppicationID As Long, Mobile As String, Message As String, Tag As String) As String Implements ISwitchPayAPI.ReceiveDeliveryReceipt
        Dim result As String = String.Empty
        Dim splits As String()
        Dim Key As String
        Dim Success As Boolean
        Dim DeliveryDate As Date
        Dim AppID As Long
        Try
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Dim a As Audit
            Try
                splits = Tag.Split("[")
                Key = splits(0)
                If Message = "0" Then
                    Success = False
                Else
                    Success = True
                End If
                DeliveryDate = CDate(splits(1))
                If AppicationID = 0 Then
                    a = (From q In db.Audits Where q.Name = CStr(Key) Select q).First()
                    AppID = a.ApplicationID
                Else
                    AppID = AppicationID
                    a = (From q In db.Audits Where (q.Name = CStr(Key)) And (q.ApplicationID = CLng(AppID)) Select q).First()
                End If
            Catch ex As Exception
                Success = False
                DeliveryDate = Now
#Disable Warning BC42104 ' Variable 'Key' is used before it has been assigned a value. A null reference exception could result at runtime.
                a = (From q In db.Audits Where q.Name = CStr(Key) Select q).First()
#Enable Warning BC42104 ' Variable 'Key' is used before it has been assigned a value. A null reference exception could result at runtime.
                AppID = a.ApplicationID
            End Try
            CreateAuditItemDetail(AppID, "SMS: " & Message & " - " & IIf(Success, "Delivered", "Failed"), 59, Key, String.Empty, False)
            result = "Success"
        Catch ex As Exception
            result = "Error: " & ex.Message
        End Try
        Return result
    End Function

    Public Function ReceiveNIUSSDMessage(AppicationID As Long, Mobile As String, Message As String, Tag As String) As String Implements ISwitchPayAPI.ReceiveNIUSSDMessage
        Dim result As String = String.Empty
        Try
            Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
            Dim splits = Tag.Split("[")
            Select Case CInt(splits(0))
                Case 6
                    Dim db As New DBDataContext(My.Settings.SwitchPayDB)
                    Dim App = (From q In db.Applications Where q.ID = CLng(AppicationID) Select q).First()
                    Dim textstr As String
                    Dim merch = (From q In rdb.Merchants Where q.ID = CLng(App.MerchantID) Select q).First()

                    If splits(1) = "1" Then
                        textstr = "Store: " & merch.Title & vbCrLf & "Application Amount: R " + String.Format("{0:0.00}", App.FinanceAmount) & String.Empty & vbCrLf & "1) Agree and Proceed" & vbCrLf & "2) Cancel"
                    Else
                        textstr = "Store: " & merch.Title & vbCrLf & vbCrLf & "OFFER" & vbCrLf & "Loan Amount: R " & String.Format("{0:0.00}", App.OfferAmount) & String.Empty & vbCrLf & "Term: " & App.OfferTerm & " Months" & vbCrLf & "Installment: R " & String.Format("{0:0.00}", App.OfferInstallment) & "/Month" & vbCrLf & String.Empty & vbCrLf & "1) Accept" & vbCrLf & "2) Decline"
                    End If
                    CreateAuditItemDetail(AppicationID, "Sent NI-USSD: " & textstr, 17, App.MobileNumber, String.Empty, False)
                    result = "text=" & textstr & "&session=1"
                Case 2
                    Dim db As New DBDataContext(My.Settings.SwitchPayDB)

                    'Dim Application As ApplicationObject = New ApplicationObject(My.Settings.WFApp)
                    'Dim Workflow As WorkflowObject

                    'If splits(1) = "1" Then
                    '	Workflow = New WorkflowObject("Authorise", Application)
                    'Else
                    '	Workflow = New WorkflowObject("Provisional Offer", Application)
                    'End If
                    'Dim WorkItemCollection As WorkItemCollection

                    'WorkItemCollection = New WorkItemCollection(Application, Workflow, New Actor(Application, "skeltalist::" & My.Settings.IntegrationUser), False)
                    'WorkItemCollection.GetRecords()

                    'For Each WorkItem In WorkItemCollection.Items
                    '	Try
                    '		If WorkItem.Subject = AppicationID Then
                    '			Try
                    '				WorkItem.CurrentContext.Variables("SMSResponse").Value = Message.ToUpper()
                    '				WorkItem.CurrentContext.SaveVariables()
                    '			Catch
                    '			End Try
                    '			If Message = "1" Then
                    '				WorkItem.Submit("Approved", String.Empty, String.Empty)
                    '			Else
                    '				WorkItem.Submit("Rejected", String.Empty, String.Empty)
                    '			End If
                    '			If splits(1) = "1" Then
                    '				CreateAuditItem(AppicationID, "NI-USSD Auth Reply Received: " & Message, 58)
                    '			Else
                    '				CreateAuditItem(AppicationID, "NI-USSD Offer Reply Received: " & Message, 58)
                    '			End If
                    '			Throw New Exception("Exit Loop")
                    '		End If
                    '	Catch
                    '	Finally
                    '		WorkItem.Dispose()
                    '	End Try
                    'Next
                    result = "text=Thank you for your response&session=0"
                Case 7
                    CreateAuditItem(AppicationID, "NI-USSD Reply Received: " & Message, 58, False)
                    result = "text=Thank you for your response&session=0"
                Case Else
                    CreateAuditItem(AppicationID, "NI-USSD Error Reply Received: " & Message, 58, False)
                    result = "text=Unfortunately an error has occured&session=0"
            End Select
        Catch ex As Exception
            CreateAuditItem(AppicationID, "NI-USSD Error Reply Received: " & Message, 58, False)
            result = "text=Unfortunately an error has occured: " & ex.Message.ToString() & "&session=0"
        End Try
        Return result
    End Function

    Public Function ReceiveSMSMessage(AppicationID As Long, Mobile As String, Message As String, Tag As String) As String Implements ISwitchPayAPI.ReceiveSMSMessage
        Dim result As String = String.Empty
        Try
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Dim Apps
            If AppicationID = 0 Then
                Apps = (From q In db.vApplications Where q.MobileNumber = Mobile Select q Order By q.ID Descending).First()
            Else
                Apps = (From q In db.vApplications Where q.ID = CLng(AppicationID) Select q Order By q.ID Descending).First()
            End If
            Dim ls As New List(Of String)
            '  ls.Add("Authorise")
            ls.Add("Provisional Offer")
            ' ls.Add("Second Phase Complete")
            Dim success As Boolean
            Dim dd As New DataDictionary(Apps.ID, My.Settings.Environment, IIf(My.Settings.Environment = "Production", "", My.Settings.Environment) & Apps.Bank, My.Settings.Environment, IIf(My.Settings.Environment = "Production", "", My.Settings.Environment) & Apps.Bank)
            Try
                dd.ApplicationD.CreateAuditItem("SMSReply Received: " & Message, 56, False)
            Catch
            End Try
            If Message.ToUpper().Contains("YES") Or (Message.ToUpper() = "Y") Then
                success = dd.WorkflowD.ActionWorkItem("Approved", ls)
            ElseIf Message.ToUpper().Contains("NO") Or (Message.ToUpper() = "N") Then
                success = dd.WorkflowD.ActionWorkItem("Rejected", ls)
            Else
                SendSMS(Apps.ID, "We did not understand your answer, please try again. Please answer YES or NO, to prevent the deal from terminating")
            End If
            result = "Success"
        Catch ex As Exception
            result = "Error: " & ex.Message
        End Try
        Return result
    End Function

    Public Function ReceiveUSSDMessage(AppicationID As Long, Mobile As String, Message As String, Tag As String) As String Implements ISwitchPayAPI.ReceiveUSSDMessage
        Dim result As String = String.Empty
        Try

        Catch ex As Exception
            result = "Error: " & ex.Message
        End Try
        Return result
    End Function

    Public Function ReceiveVodacomDeliveryReceipt(AppicationID As Long, Mobile As String, Message As String, Tag As String) As String Implements ISwitchPayAPI.ReceiveVodacomDeliveryReceipt
        Dim result As String = String.Empty
        Try

        Catch ex As Exception
            result = "Error: " & ex.Message
        End Try
        Return result
    End Function

    Public Function ReceiveVodacomNIUSSDMessage(AppicationID As Long, Mobile As String, Message As String, Tag As String) As String Implements ISwitchPayAPI.ReceiveVodacomNIUSSDMessage
        Dim result As String = String.Empty
        Try

        Catch ex As Exception
            result = "Error: " & ex.Message
        End Try
        Return result
    End Function

    Public Function ReceiveVodacomSMSMessage(AppicationID As Long, Mobile As String, Message As String, Tag As String) As String Implements ISwitchPayAPI.ReceiveVodacomSMSMessage
        Dim result As String = String.Empty
        Try

        Catch ex As Exception
            result = "Error: " & ex.Message
        End Try
        Return result
    End Function

    Public Function ReceiveVodacomUSSDMessage(AppicationID As Long, Mobile As String, Message As String, Tag As String) As String Implements ISwitchPayAPI.ReceiveVodacomUSSDMessage
        Dim result As String = String.Empty
        Try

        Catch ex As Exception
            result = "Error: " & ex.Message
        End Try
        Return result
    End Function

    Function RedeemApplication(MerchantID As Long, TerminalID As String, FinanceAmount As Double, GenerateOTP As Boolean, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.RedeemApplication
        Dim rxsd As New SwitchPayIntegration.Models.Response
        Dim hasMerchant As Boolean = False
        Dim hasIDNumber As Boolean = False
        Dim hasMobile As Boolean = False
        Dim nofilters As Boolean = False
        Dim results As Boolean = True
        Dim amountvalid As Boolean = False
        Dim bankvalid As Boolean = False
        Dim mobilevalid As Boolean = True
        Dim idvalid As Boolean = True
        Dim status As Long = 0
        Dim AmountTooBig As Boolean = False
        If (Not IsNumeric(IDNumber)) And (IDNumber <> String.Empty) Then
            idvalid = False
        End If
        If (Not IsNumeric(MobileNumber)) And (MobileNumber <> String.Empty) Then
            mobilevalid = False
        End If
        Try
            If (ApplicationID = 0) And (ApplicationRef = String.Empty) And (IDNumber = String.Empty) And (MobileNumber = String.Empty) Then
                nofilters = True
            End If


            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
            Try
                If ApplicationID <> 0 Then
                    App = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
                    status = App.AuditTypeID
                ElseIf IDNumber <> String.Empty Then
                    App = (From q In db.Applications Where q.IDNumber = CStr(IDNumber) Select q Order By q.DateCreated Descending).First()
                    status = App.AuditTypeID
                ElseIf ApplicationRef <> String.Empty Then
                    App = (From q In db.Applications Where q.Reference = CStr(ApplicationRef) Select q Order By q.DateCreated Descending).First()
                    status = App.AuditTypeID
                Else
                    'Apps = (From q In db.Applications Where q.entityi = CStr(ApplicationRef) Select q).ToArray
                    results = False
                End If
                If FinanceAmount <> App.OfferAmount Then
                    AmountTooBig = True
                End If
                If (FinanceAmount >= 500) And (FinanceAmount <= 230000) And (Not AmountTooBig) Then
                    amountvalid = True
                End If
            Catch
                results = False
            End Try
            hdr = hdt.NewHeaderRow()
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim fdr111 As SwitchPayIntegration.Models.Response.FieldsRow
            fdr111 = fdt.NewFieldsRow()
            fdr111("Name") = "Collectable"
            If CanCollect(App.ID) Then
                fdr111("Value") = True
            Else
                fdr111("Value") = False
            End If
            fdt.AddFieldsRow(fdr111)

            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If idvalid And mobilevalid And amountvalid And CanCollect(App.ID) And (Not nofilters) Then
                hdr("ApplicationID") = App.ID
                hdr("MerchantID") = MerchantID
                Merchant = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
                hasMerchant = True
                hdr("ApplicationRef") = App.Reference
                hdr("MerchantRef") = Merchant.Reference
                hdr("IDNumber") = App.IDNumber
                hdr("MobileNumber") = App.MobileNumber
                hdr("IsError") = False
                Dim fdr5 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr5 = fdt.NewFieldsRow()
                fdr5("Name") = "Screen Message"
                fdr5("Value") = "Redeemed"
                fdt.AddFieldsRow(fdr5)
                Dim fdr3 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr3 = fdt.NewFieldsRow()
                fdr3("Name") = "SlipMessage"
                fdr3("Value") = "Application Redeemed"
                fdt.AddFieldsRow(fdr3)
                Dim fdr4 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr4 = fdt.NewFieldsRow()
                fdr4("Name") = "TransactionType"
                fdr4("Value") = "Redemption"
                fdt.AddFieldsRow(fdr4)
                Dim fdr6 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr6 = fdt.NewFieldsRow()
                fdr6("Name") = "AvailableBalance"
                fdr6("Value") = Math.Round(CDbl(App.OfferAmount), 2).ToString().Replace(",", ".")
                fdt.AddFieldsRow(fdr6)
                Dim fdr7 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr7 = fdt.NewFieldsRow()
                fdr7("Name") = "BalanceAfterRedeem"
                fdr7("Value") = Math.Round(CDbl(App.OfferAmount - FinanceAmount), 2)
                fdt.AddFieldsRow(fdr7)
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "0"
                mdr2("Message") = "Application ID: " & App.ID & " redeemed successfully."
                mdr2("IsError") = False
                mdt.AddMessagesRow(mdr2)
            Else
                hdr("IsError") = True
            End If
            hdt.AddHeaderRow(hdr)
#Disable Warning BC42024 ' Unused local variable: 'mdr'.
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
#Enable Warning BC42024 ' Unused local variable: 'mdr'.
            If nofilters Then
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "3"
                mdr2("Message") = "No Filters Provided."
                mdr2("IsError") = True
                mdt.AddMessagesRow(mdr2)
            End If
            If Not results Then
                Dim mdr4 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr4 = mdt.NewMessagesRow()
                mdr4("Code") = "2"
                mdr4("Message") = "No Applications Found."
                mdr4("IsError") = True
                mdt.AddMessagesRow(mdr4)
            End If
            If CanCollect(App.ID) Then
            Else
                Dim mdr7 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr7 = mdt.NewMessagesRow()
                mdr7("Code") = "2"
                mdr7("Message") = "Not Ready For Collection."
                mdr7("IsError") = True
                mdt.AddMessagesRow(mdr7)
            End If
            If status = 25 Then
                Dim mdr7 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr7 = mdt.NewMessagesRow()
                mdr7("Code") = "6"
                mdr7("Message") = "AlReady Been Collected."
                mdr7("IsError") = True
                mdt.AddMessagesRow(mdr7)
            End If
            If Not amountvalid Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "5"
                mdr3("Message") = "Amount Invalid"
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            If Not mobilevalid Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "5"
                mdr3("Message") = "Mobile Number Invalid"
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            If Not idvalid Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "5"
                mdr3("Message") = "ID Number Invalid"
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            If GenerateOTP Then
                'Dim url2 = "https://www.xml2sms.gsm.co.za/send/?username=warpdev&password=Vodacom963&number=" & MobileNumber & "&message=Welcome to Switch Pay, your payment OTP is 11111"
                'Dim client2 As New WebClient
                'Dim Xml2 = client2.DownloadString(url2)
            End If

        Catch ex As Exception
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            'Dim apps = (From q In db.Applications Where appl)
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            hdr = hdt.NewHeaderRow()
            hdr("IsError") = True
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If hasMerchant = False Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "4"
                mdr3("Message") = "No Valid Merchant Found."
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "1"
            mdr("Message") = "Exception: " & ex.Message
            mdr("IsError") = True
            mdt.AddMessagesRow(mdr)
        End Try

        'Dim url = "https://www.xml2sms.gsm.co.za/send/?username=warpdev&password=Vodacom963&number=" & MobileNumber & "&message=Welcome to Switch Pay, your payment OTP is 11111"
        'Dim client As New WebClient
        'Dim Xml = client.DownloadString(url)

        Return rxsd
    End Function

    Function RegisterMerchant(MerchantRef As String, Name As String, ContactNumber As String, URL As String, Email As String, IndustryID As Long) As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.RegisterMerchant
        Dim rxsd As New SwitchPayIntegration.Models.Response
        Dim hasMerchant As Boolean = True
        Dim contactvalid As Boolean = True
        Try
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            If Name = String.Empty Then
                hasMerchant = False

            End If
            If (ContactNumber = String.Empty) Or (Not IsNumeric(ContactNumber)) Then
                contactvalid = False
            End If

            If hasMerchant And contactvalid Then
                LoadMerchantDetailsData(0)
                Merchant.Title = Name
                'Merchant.URL = URL
                'Merchant.EntityTypeID = 2
                'Merchant.Type = "Company"
                'Merchant.IsDeleted = False
                'Merchant.IndustryID = IndustryID
                'merchantcellphone.Phone.Title = ContactNumber
                'merchantcompany.Title = Name
                'merchantcompany.WebsiteAdress = URL
                'merchantcompany.IsDeleted = False
                'Merchant.Reference = MerchantRef
                'Dim MerchantHistory As New EntityHistory
                'MerchantHistory.Title = "New Merchant created."
                'MerchantHistory.DateCreated = Now
                'MerchantHistory.Entity = Merchant
                'rdb.Merchants.InsertOnSubmit(Merchant)
                'db.SubmitChanges()
                'Merchant.Company.Title = Name
                'Merchant.Company.IsDeleted = False
                'Merchant.Company.WebsiteAdress = URL
                'db.SubmitChanges()
                Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
                hdt = rxsd.Tables("Header")
                Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
                hdr = hdt.NewHeaderRow()
                hdr("MerchantID") = Merchant.ID
                hdr("MerchantRef") = MerchantRef
                hdt.AddHeaderRow(hdr)
                Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
                fdt = rxsd.Tables("Fields")
                Dim fdr As SwitchPayIntegration.Models.Response.FieldsRow
                fdr = fdt.NewFieldsRow()
                fdr("Name") = "MerchantID"
                fdr("Value") = Merchant.ID
                fdt.AddFieldsRow(fdr)

                Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
                mdt = rxsd.Tables("Messages")
                Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
                mdr = mdt.NewMessagesRow()
                mdr("Code") = "0"
                mdr("Message") = "Merchant ID: " & MerchantRef & " registered"
                mdr("IsError") = False
                mdt.AddMessagesRow(mdr)
            Else
                Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
                hdt = rxsd.Tables("Header")
                Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
                hdr = hdt.NewHeaderRow()
                hdr("MerchantRef") = MerchantRef
                hdr("IsError") = True

                hdt.AddHeaderRow(hdr)
                Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
                fdt = rxsd.Tables("Fields")

                Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
                mdt = rxsd.Tables("Messages")
                If Not hasMerchant Then
                    Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                    mdr2 = mdt.NewMessagesRow()
                    mdr2("Code") = "3"
                    mdr2("Message") = "No Name Provided."
                    mdr2("IsError") = True
                    mdt.AddMessagesRow(mdr2)
                End If
                If Not contactvalid Then
                    Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                    mdr2 = mdt.NewMessagesRow()
                    mdr2("Code") = "3"
                    mdr2("Message") = "No Contact Number Provided."
                    mdr2("IsError") = True
                    mdt.AddMessagesRow(mdr2)
                End If

            End If
        Catch ex As Exception
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            'Dim apps = (From q In db.Applications Where appl)
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            hdr = hdt.NewHeaderRow()
            hdr("IsError") = True
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If hasMerchant = False Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "4"
                mdr3("Message") = "No Valid Merchant Found."
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "1"
            mdr("Message") = "Exception: " & ex.Message
            mdr("IsError") = True
            mdt.AddMessagesRow(mdr)
        End Try
        Return rxsd
    End Function

    Function RegisterMerchantSkelta(MerchantRef As String, Name As String, ContactNumber As String, URL As String, Email As String, IndustryID As Long) As String Implements ISwitchPayAPI.RegisterMerchantSkelta
        Dim rxsd As New SwitchPayIntegration.Models.Response
        Dim hasMerchant As Boolean = True
        Dim contactvalid As Boolean = True
        Try
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            If Name = String.Empty Then
                hasMerchant = False

            End If
            If (ContactNumber = String.Empty) Or (Not IsNumeric(ContactNumber)) Then
                contactvalid = False
            End If

            If hasMerchant And contactvalid Then
                LoadMerchantDetailsData(0)
                'Merchant.Title = Name
                'Merchant.URL = URL
                'Merchant.EntityTypeID = 2
                'Merchant.Type = "Company"
                'Merchant.IsDeleted = False
                'Merchant.IndustryID = IndustryID
                'merchantcellphone.Phone.Title = ContactNumber
                'merchantcompany.Title = Name
                'merchantcompany.WebsiteAdress = URL
                'merchantcompany.IsDeleted = False
                'Merchant.Reference = MerchantRef
                'Dim MerchantHistory As New EntityHistory
                'MerchantHistory.Title = "New Merchant created."
                'MerchantHistory.DateCreated = Now
                'MerchantHistory.Entity = Merchant
                'rdb.Merchants.InsertOnSubmit(Merchant)
                db.SubmitChanges()
                Merchant.Title = Name
                Merchant.IsDeleted = False
                ' Merchant.MerchantDetails.FirstOrDefault. = URL
                db.SubmitChanges()
                Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
                hdt = rxsd.Tables("Header")
                Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
                hdr = hdt.NewHeaderRow()
                hdr("MerchantID") = Merchant.ID
                hdr("MerchantRef") = MerchantRef
                hdt.AddHeaderRow(hdr)
                Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
                fdt = rxsd.Tables("Fields")
                Dim fdr As SwitchPayIntegration.Models.Response.FieldsRow
                fdr = fdt.NewFieldsRow()
                fdr("Name") = "MerchantID"
                fdr("Value") = Merchant.ID
                fdt.AddFieldsRow(fdr)

                Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
                mdt = rxsd.Tables("Messages")
                Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
                mdr = mdt.NewMessagesRow()
                mdr("Code") = "0"
                mdr("Message") = "Merchant ID: " & MerchantRef & " registered"
                mdr("IsError") = False
                mdt.AddMessagesRow(mdr)
            Else
                Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
                hdt = rxsd.Tables("Header")
                Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
                hdr = hdt.NewHeaderRow()
                hdr("MerchantRef") = MerchantRef
                hdr("IsError") = True

                hdt.AddHeaderRow(hdr)
                Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
                fdt = rxsd.Tables("Fields")

                Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
                mdt = rxsd.Tables("Messages")
                If Not hasMerchant Then
                    Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                    mdr2 = mdt.NewMessagesRow()
                    mdr2("Code") = "3"
                    mdr2("Message") = "No Name Provided."
                    mdr2("IsError") = True
                    mdt.AddMessagesRow(mdr2)
                End If
                If Not contactvalid Then
                    Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                    mdr2 = mdt.NewMessagesRow()
                    mdr2("Code") = "3"
                    mdr2("Message") = "No Contact Number Provided."
                    mdr2("IsError") = True
                    mdt.AddMessagesRow(mdr2)
                End If

            End If
        Catch ex As Exception
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            'Dim apps = (From q In db.Applications Where appl)
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            hdr = hdt.NewHeaderRow()
            hdr("IsError") = True
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If hasMerchant = False Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "4"
                mdr3("Message") = "No Valid Merchant Found."
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "1"
            mdr("Message") = "Exception: " & ex.Message
            mdr("IsError") = True
            mdt.AddMessagesRow(mdr)
        End Try
        Return Merchant.ID.ToString()


    End Function


    Public Sub ReleaseCommFile(FileID As Long) Implements ISwitchPayAPI.ReleaseCommFile
        Dim DataD As New DataDictionary(My.Settings.Environment, My.Settings.Repository)
        Dim p As New PaymentDictionary(DataD, "c:\templates\")
        p.ReleaseCommFile(FileID)
    End Sub

    Public Sub ReleaseFile(FileID As Long) Implements ISwitchPayAPI.ReleaseFile
        Dim DataD As New DataDictionary(My.Settings.Environment, My.Settings.Repository)
        Dim p As New PaymentDictionary(DataD, "c:\templates\")
        p.ReleaseFile(FileID)
    End Sub

    Function ReturnApplication(MerchantID As Long, TerminalID As String, FinanceAmount As Double, GenerateOTP As Boolean, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.ReturnApplication
        Dim rxsd As New SwitchPayIntegration.Models.Response
        Dim hasMerchant As Boolean = False
        Try
            Dim nofilters As Boolean = False
            Dim results As Boolean = True
            If (ApplicationID = 0) And (ApplicationRef = String.Empty) And (IDNumber = String.Empty) And (MobileNumber = String.Empty) Then
                nofilters = True
            End If
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            If ApplicationID <> 0 Then
                App = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
            ElseIf ApplicationRef <> String.Empty Then
                App = (From q In db.Applications Where q.Reference = CStr(ApplicationRef) Select q Order By q.DateCreated Descending).First()
            ElseIf IDNumber <> String.Empty Then
                App = (From q In db.Applications Where q.IDNumber = CStr(IDNumber) Select q Order By q.DateCreated Descending).First()
            Else
                'Apps = (From q In db.Applications Where q.entityi = CStr(ApplicationRef) Select q).ToArray
                results = False
            End If

            hdr = hdt.NewHeaderRow()
            Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
            If results Then
                hdr("ApplicationID") = App.ID
                hdr("MerchantID") = MerchantID
                Merchant = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
                hasMerchant = True
                hdr("ApplicationRef") = App.Reference
                hdr("MerchantRef") = Merchant.Reference
                hdr("IDNumber") = App.IDNumber
                hdr("MobileNumber") = App.MobileNumber
                hdr("IsError") = False
            Else
                hdr("IsError") = True
            End If
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")

            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
#Disable Warning BC42024 ' Unused local variable: 'mdr'.
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
#Enable Warning BC42024 ' Unused local variable: 'mdr'.
            If nofilters Then
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "3"
                mdr2("Message") = "No Filters Provided."
                mdr2("IsError") = True
                mdt.AddMessagesRow(mdr2)
            ElseIf Not results Then
                Dim mdr4 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr4 = mdt.NewMessagesRow()
                mdr4("Code") = "2"
                mdr4("Message") = "No Applications Found."
                mdr4("IsError") = True
                mdt.AddMessagesRow(mdr4)
            Else
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "0"
                mdr2("Message") = "Application ID: " & App.ID & " returned"
                mdr2("IsError") = False
                mdt.AddMessagesRow(mdr2)
            End If
            If GenerateOTP Then
                Dim url = "https://www.xml2sms.gsm.co.za/send/?username=warpdev&password=Vodacom963&number=" & MobileNumber & "&message=Welcome to Switch Pay, your OTP is 11111"
                Dim client As New WebClient
                Dim Xml = client.DownloadString(url)
            End If

        Catch ex As Exception
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            'Dim apps = (From q In db.Applications Where appl)
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            hdr = hdt.NewHeaderRow()
            hdr("IsError") = True
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If hasMerchant = False Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "4"
                mdr3("Message") = "No Valid Merchant Found."
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "1"
            mdr("Message") = "Exception: " & ex.Message
            mdr("IsError") = True
            mdt.AddMessagesRow(mdr)
        End Try
        Return rxsd
    End Function

    Public Sub SaveDisbursementFileItems(FileID As Long, Items As String) Implements ISwitchPayAPI.SaveDisbursementFileItems
        Dim DataD As New DataDictionary(My.Settings.Environment, My.Settings.Repository)
        Dim p As New PaymentDictionary(DataD, "c:\templates\")
        p.SaveDisbursementFileItems(FileID, Items)
    End Sub

    Public Sub SaveDOFileItems(FileID As Long, Items As String) Implements ISwitchPayAPI.SaveDOFileItems
        Dim DataD As New DataDictionary(My.Settings.Environment, My.Settings.Repository)
        Dim p As New PaymentDictionary(DataD, "c:\templates\")
        p.SaveDOFileItems(FileID, Items)
    End Sub

    Function SendApplicationOTP(MerchantID As Long, TerminalID As String, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.SendApplicationOTP
        Dim rxsd As New SwitchPayIntegration.Models.Response
        Dim hasMerchant As Boolean = False
        Dim mobilevalid As Boolean = True
        Dim idvalid As Boolean = True
        If (Not IsNumeric(IDNumber)) And (IDNumber <> String.Empty) Then
            idvalid = False
        End If
        If (Not IsNumeric(MobileNumber)) And (MobileNumber <> String.Empty) Then
            mobilevalid = False
        End If
        Try
            Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
            Dim nofilters As Boolean = False
            Dim results As Boolean = True
            If (ApplicationID = 0) And (ApplicationRef = String.Empty) And (IDNumber = String.Empty) And (MobileNumber = String.Empty) Then
                nofilters = True
            End If
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Merchant = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
            hasMerchant = True
            Try
                If ApplicationID <> 0 Then
                    App = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
                ElseIf ApplicationRef <> String.Empty Then
                    App = (From q In db.Applications Where q.Reference = CStr(ApplicationRef) Select q Order By q.DateCreated Descending).First()
                ElseIf IDNumber <> String.Empty Then
                    App = (From q In db.Applications Where q.IDNumber = CStr(IDNumber) Select q Order By q.DateCreated Descending).First()
                Else
                    'Apps = (From q In db.Applications Where q.entityi = CStr(ApplicationRef) Select q).ToArray
                    results = False
                End If
            Catch
                results = False
            End Try
            hdr = hdt.NewHeaderRow()
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")

            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If results Then
                hdr("ApplicationID") = App.ID
                hdr("MerchantID") = MerchantID
                hdr("ApplicationRef") = App.Reference
                hdr("MerchantRef") = Merchant.Reference
                hdr("IDNumber") = App.IDNumber
                hdr("MobileNumber") = App.MobileNumber
                hdr("IsError") = False

                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "0"
                mdr2("Message") = "Application ID: " & App.ID & " OTP successfully sent"
                mdr2("IsError") = False
                mdt.AddMessagesRow(mdr2)
            Else
                hdr("IsError") = True
            End If
            hdt.AddHeaderRow(hdr)
#Disable Warning BC42024 ' Unused local variable: 'mdr'.
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
#Enable Warning BC42024 ' Unused local variable: 'mdr'.
            If nofilters Then
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "3"
                mdr2("Message") = "No Filters Provided."
                mdr2("IsError") = True
                mdt.AddMessagesRow(mdr2)
            End If
            If Not results Then
                Dim mdr4 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr4 = mdt.NewMessagesRow()
                mdr4("Code") = "2"
                mdr4("Message") = "No Applications Found."
                mdr4("IsError") = True
                mdt.AddMessagesRow(mdr4)
            End If
            If Not idvalid Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "5"
                mdr3("Message") = "ID Number Invalid"
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            If Not mobilevalid Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "5"
                mdr3("Message") = "Mobile Number Invalid"
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            'Dim url = "https://www.xml2sms.gsm.co.za/send/?username=warpdev&password=Vodacom963&number=" & MobileNumber & "&message=Welcome to Switch Pay, your OTP is 11111"
            'Dim client As New WebClient
            'Dim Xml = client.DownloadString(url)

        Catch ex As Exception
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            'Dim apps = (From q In db.Applications Where appl)
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            hdr = hdt.NewHeaderRow()
            hdr("IsError") = True
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If hasMerchant = False Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "4"
                mdr3("Message") = "No Valid Merchant Found."
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "1"
            mdr("Message") = "Exception: " & ex.Message
            mdr("IsError") = True
            mdt.AddMessagesRow(mdr)
        End Try

        Return rxsd
    End Function

    Public Function SendAuthSMS(ApplicationID As Long) As String Implements ISwitchPayAPI.SendAuthSMS
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
        Dim app = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
        Dim merch = (From q In rdb.Merchants Where q.ID = CLng(app.MerchantID) Select q).First()
        'Dim ser As JObject = JObject.Parse(Response)
        Dim tmobile As String = "27" & app.MobileNumber.Substring(1, app.MobileNumber.Length - 1)
        Dim charArr As Char() = "0123456789".ToCharArray()
        Dim strrandom As String = String.Empty
        Dim objran As New Random
        Dim noofcharacters As Integer = 5
        For i As Integer = 0 To noofcharacters - 1
            'It will not allow Repetation of Characters
            Dim pos As Integer = objran.[Next](1, charArr.Length)
            If Not strrandom.Contains(charArr.GetValue(pos).ToString()) Then
                strrandom += charArr.GetValue(pos)
            Else
                i -= 1
            End If
        Next
        app.AcceptPIN = strrandom
        db.SubmitChanges()
        Dim mer = (From q In rdb.Merchants Where q.ID = CStr(app.MerchantID) Select q).First()

        Dim resp As String
        Dim client As New WebClient

        resp = "Please enter apply code " & strrandom & " to approve Fin App " & app.ID & " at " & mer.ShortName & " and to consent to credit check with Credit Bureaus. Queries 0861995008."



        Return SendSMSVodacom(ApplicationID, resp)
    End Function

    Public Function SendBankMail(ApplicationID As Long, BankID As Long) As String Implements ISwitchPayAPI.SendBankMail
        '        Dim message As String = String.Empty
        '        Try
        '            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        '            Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
        '            Dim Apps = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
        '            Dim merch = (From q In rdb.Merchants Where q.ID = CLng(Apps.MerchantID) Select q).First()
        '            Dim email2 As New MailMessage

        '            email2.From = New MailAddress("workflow@switchpay.co.za")
        '            email2.ReplyToList.Add(New MailAddress("cpresponses@switchpay.co.za"))
        '            Dim SMTP As New SmtpClient("smtp.gmail.com")
        '            SMTP.UseDefaultCredentials = False
        '            SMTP.Credentials = New System.Net.NetworkCredential("workflow@switchpay.co.za", "selfadrpcbiajyux") '<-- Password Here
        '            SMTP.EnableSsl = True
        '            email2.Subject = "SwitchPay Customer Application " & ApplicationID & " - Customer ID " & Apps.ID
        '            email2.To.Add("bryan@switchpay.co.za")
        '            email2.To.Add("admin@switchpay.co.za")

        '            If BankID = 6 Then
        '                If (Not My.Settings.Environment = "Production") And (merch.ID <> 9) And (merch.ID <> 1) Then
        '                    email2.To.Add("retailers@capitecbank.co.za")
        '                    email2.To.Add("ginogoodall@capitecbank.co.za")
        '                End If
        '                Dim term, BEE, employer, employmenttype, position, dateemployed, payfreq As String
        '                term = (From q In db.ApplicationFieldValues Where (q.ApplicationID = CLng(ApplicationID)) And (q.FieldDefinitionEntityID = 81) Select q.Title).First()
        '                BEE = (From q In db.ApplicationFieldValues Where (q.ApplicationID = CLng(ApplicationID)) And (q.FieldDefinitionEntityID = 83) Select q.Title).First()
        '                employer = (From q In db.ApplicationFieldValues Where (q.ApplicationID = CLng(ApplicationID)) And (q.FieldDefinitionEntityID = 64) Select q.Title).First()
        '                employmenttype = (From q In db.ApplicationFieldValues Where (q.ApplicationID = CLng(ApplicationID)) And (q.FieldDefinitionEntityID = 68) Select q.Title).First()
        '                position = (From q In db.ApplicationFieldValues Where (q.ApplicationID = CLng(ApplicationID)) And (q.FieldDefinitionEntityID = 66) Select q.Title).First()
        '                dateemployed = (From q In db.ApplicationFieldValues Where (q.ApplicationID = CLng(ApplicationID)) And (q.FieldDefinitionEntityID = 67) Select q.Title).First()
        '                payfreq = (From q In db.ApplicationFieldValues Where (q.ApplicationID = CLng(ApplicationID)) And (q.FieldDefinitionEntityID = 69) Select q.Title).First()

        '                message = "<table border=""0"" cellspacing=""0"" cellpadding=""0"">
        '    <tbody>
        '        <tr>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: Left(); background-color: transparent;"">
        '            <p><strong><span style=""font-family Calibri;"">APPLICATION INFO:</span></strong></p>
        '            </td>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: Left(); background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: Left(); background-color: transparent;"">
        '            <p><span style=""font-family Calibri;"">Application Reference Number: </span></p>
        '            </td>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: Left(); background-color: transparent;"">
        '            <p><span style=""font-family Calibri;"">" & ApplicationID & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: Left(); background-color: transparent;"">
        '            <p><span style=""font-family Calibri;"">Date:</span></p>
        '            </td>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: Left(); background-color: transparent;"">
        '            <p><span style=""font-family Calibri;"">" & CDate(Apps.DateCreated).ToLongDateString() & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: Left(); background-color: transparent;"">
        '            <p><span style=""font-family Calibri;"">Time: </span></p>
        '            </td>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: Left(); background-color: transparent;"">
        '            <p><span style=""font-family Calibri;"">" & CDate(Apps.DateCreated).ToShortTimeString() & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: Left(); background-color: transparent;"">
        '            <p><span style=""font-family Calibri;"">Status: </span></p>
        '            </td>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: Left(); background-color: transparent;"">
        '            <p><span style=""font-family Calibri;"">" & Apps.AuditTypeID & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: Left(); background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: Left(); background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: Left(); background-color: transparent;"">
        '            <p><strong><span style=""font-family Calibri;"">PRODUCT LOAN INFORMATION:</span></strong></p>
        '            </td>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: Left(); background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: Left(); background-color: transparent;"">
        '            <p><span style=""font-family Calibri;"">Amount:</span></p>
        '            </td>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: Left(); background-color: transparent;"">
        '            <p><span style=""font-family Calibri;"">" & String.Format("{0:C2}", Apps.FinanceAmount) & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: Left(); background-color: transparent;"">
        '            <p><span style=""font-family Calibri;"">Requested Term:</span></p>
        '            </td>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: Left(); background-color: transparent;"">
        '            <p><span style=""font-family Calibri;"">" & term & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: Left(); background-color: transparent;"">
        '            <p><span style=""font-family Calibri;"">Merchant:</span></p>
        '            </td>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: Left(); background-color: transparent;"">
        '            <p><span style=""font-family Calibri;"">" & merch.Title & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: Left(); background-color: transparent;"">
        '            <p><span style=""font-family Calibri;"">Product Category:</span></p>
        '            </td>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: Left(); background-color: transparent;"">
        '            <p><span style=""font-family Calibri;"">" & merch.Title & "</span></p>
        '            </td>
        '        </tr>

        '        <tr>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: Left(); background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: Left(); background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: Left(); background-color: transparent;"">
        '            <p><strong><span style=""font-family Calibri;"">CUSTOMER PERSONAL INFORMATION: </span></strong></p>
        '            </td>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: Left(); background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: Left(); background-color: transparent;"">
        '            <p><span style=""font-family Calibri;"">First Name:</span></p>
        '            </td>
        '            <td valign=""top"" style=""padding 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & Apps.FirstName & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Surname:</span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & Apps.Surname & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">ID Number: </span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & Apps.IDNumber & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Mobile Number: </span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & Apps.MobileNumber & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: left(); background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: left(); background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: left(); background-color: transparent;"">
        '            <p><strong><span style=""font-family: Calibri;"">INCOME AND EMPLOYMENT INFORMATION:</span></strong></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: left(); background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Salary after deductions: </span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & String.Format("{0:C2}", Apps.NettIncome) & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Employer Name:</span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & employer & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Employment Type:</span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & employmenttype & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Position at work:</span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & position & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Started Working From:</span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & dateemployed & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Payment Frequency:</span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & payfreq & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: left(); background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: left(); background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: left(); background-color: transparent;"">
        '            <p><strong><span style=""font-family: Calibri;"">TERMS AND CONDITIONS</span></strong></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: left(); background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">I am historically disadvantaged: </span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & BEE & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Consent to use my personal details to do a credit bureau enquiry:</span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">True</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Accepted Provider terms of use: </span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">True</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 311.6pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Accepted SwitchPay terms and conditions: </span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 139.2pt; text-align: left(); background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">True</span></p>
        '            </td>
        '        </tr>
        '    </tbody>
        '</table>"
        '            ElseIf BankID = 5 Then
        '                If Not My.Settings.WFApp.Contains("UAT") Then
        '                    email2.To.Add("switchpay@loanzie.co.za")
        '                End If
        '                Dim resstatus, nextdate, pcode, address, email, marital, ethnic, lang, city, hrcontact, employer, employmenttype, position, dateemployed, payfreq As String
        '                'resstatus = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 56 Select q.Title).First
        '                'nextdate = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 50 Select q.Title).First
        '                'address = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 57 Select q.Title).First
        '                'email = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 45 Select q.Title).First
        '                'marital = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 53 Select q.Title).First
        '                'ethnic = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 54 Select q.Title).First
        '                'lang = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 55 Select q.Title).First
        '                'city = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 58 Select q.Title).First
        '                'pcode = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 59 Select q.Title).First
        '                'hrcontact = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 60 Select q.Title).First
        '                'employer = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 44 Select q.Title).First
        '                'employmenttype = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 48 Select q.Title).First
        '                'position = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 46 Select q.Title).First
        '                'dateemployed = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 47 Select q.Title).First
        '                'payfreq = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 49 Select q.Title).First

        '                message = "<table border=""0"" cellspacing=""0"" cellpadding=""0"">
        '                    <tbody>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><strong><span style=""font-family: Calibri;"">APPLICATION INFO:</span></strong></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p>&nbsp;</p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">Application Reference Number: </span></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">" & ApplicationID & "</span></p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">Date:</span></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">" & CDate(Apps.DateCreated).ToLongDateString & "</span></p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">Time: </span></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">" & CDate(Apps.DateCreated).ToLongDateString & "</span></p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">Status: </span></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">" & Apps.AuditType.Title & "</span></p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p>&nbsp;</p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p>&nbsp;</p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><strong><span style=""font-family: Calibri;"">PRODUCT LOAN INFORMATION:</span></strong></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p>&nbsp;</p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">Amount:</span></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">" & String.Format("{0:C2}", Apps.FinanceAmount) & "</span></p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">Requested Term:</span></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">6</span></p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">Merchant:</span></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">" & merch.Title & "</span></p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">Product Category:</span></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">" & merch.Industry.Title & "</span></p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p>&nbsp;</p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p>&nbsp;</p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><strong><span style=""font-family: Calibri;"">CUSTOMER PERSONAL INFORMATION: </span></strong></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p>&nbsp;</p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">First Name:</span></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">" & Apps.FirstName & "</span></p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">Surname:</span></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">" & Apps.Surname & "</span></p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">ID Number: </span></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">" & Apps.IDNumber & "</span></p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">Mobile Number: </span></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">" & Apps.MobileNumber & "</span></p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p>&nbsp;</p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p>&nbsp;</p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><strong><span style=""font-family: Calibri;"">TERMS AND CONDITIONS</span></strong></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p>&nbsp;</p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">Marketing Consent: </span></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">True</span></p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">Consent to use my personal details to do a credit bureau enquiry:</span></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">True</span></p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">Accepted Provider terms of use: </span></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">True</span></p>
        '                            </td>
        '                        </tr>
        '                        <tr>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">Accepted SwitchPay terms and conditions: </span></p>
        '                            </td>
        '                            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '                            <p><span style=""font-family: Calibri;"">True</span></p>
        '                            </td>
        '                        </tr>
        '                    </tbody>
        '                </table>
        '                <br />"
        '            Else
        '                If (Not My.Settings.Environment = "Production") And (merch.ID <> 9) And (merch.ID <> 1) Then
        '                    email2.To.Add("switchpay@loanzie.co.za")
        '                End If
        '#Disable Warning BC42024 ' Unused local variable: 'city'.
        '#Disable Warning BC42024 ' Unused local variable: 'lang'.
        '#Disable Warning BC42024 ' Unused local variable: 'marital'.
        '#Disable Warning BC42024 ' Unused local variable: 'payfreq'.
        '#Disable Warning BC42024 ' Unused local variable: 'dateemployed'.
        '#Disable Warning BC42024 ' Unused local variable: 'email'.
        '#Disable Warning BC42024 ' Unused local variable: 'nextdate'.
        '#Disable Warning BC42024 ' Unused local variable: 'hrcontact'.
        '#Disable Warning BC42024 ' Unused local variable: 'employmenttype'.
        '#Disable Warning BC42024 ' Unused local variable: 'employer'.
        '#Disable Warning BC42024 ' Unused local variable: 'address'.
        '#Disable Warning BC42024 ' Unused local variable: 'position'.
        '#Disable Warning BC42024 ' Unused local variable: 'resstatus'.
        '#Disable Warning BC42024 ' Unused local variable: 'ethnic'.
        '#Disable Warning BC42024 ' Unused local variable: 'pcode'.
        '                Dim resstatus, nextdate, pcode, address, email, marital, ethnic, lang, city, hrcontact, employer, employmenttype, position, dateemployed, payfreq As String
        '#Enable Warning BC42024 ' Unused local variable: 'pcode'.
        '#Enable Warning BC42024 ' Unused local variable: 'ethnic'.
        '#Enable Warning BC42024 ' Unused local variable: 'resstatus'.
        '#Enable Warning BC42024 ' Unused local variable: 'position'.
        '#Enable Warning BC42024 ' Unused local variable: 'address'.
        '#Enable Warning BC42024 ' Unused local variable: 'employer'.
        '#Enable Warning BC42024 ' Unused local variable: 'employmenttype'.
        '#Enable Warning BC42024 ' Unused local variable: 'hrcontact'.
        '#Enable Warning BC42024 ' Unused local variable: 'nextdate'.
        '#Enable Warning BC42024 ' Unused local variable: 'email'.
        '#Enable Warning BC42024 ' Unused local variable: 'dateemployed'.
        '#Enable Warning BC42024 ' Unused local variable: 'payfreq'.
        '#Enable Warning BC42024 ' Unused local variable: 'marital'.
        '#Enable Warning BC42024 ' Unused local variable: 'lang'.
        '#Enable Warning BC42024 ' Unused local variable: 'city'.
        '                resstatus = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 56 Select q.Title).First
        '                nextdate = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 50 Select q.Title).First
        '                address = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 57 Select q.Title).First
        '                email = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 45 Select q.Title).First
        '                marital = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 53 Select q.Title).First
        '                ethnic = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 54 Select q.Title).First
        '                lang = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 55 Select q.Title).First
        '                city = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 58 Select q.Title).First
        '                pcode = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 59 Select q.Title).First
        '                hrcontact = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 60 Select q.Title).First
        '                employer = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 44 Select q.Title).First
        '                employmenttype = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 48 Select q.Title).First
        '                position = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 46 Select q.Title).First
        '                dateemployed = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 47 Select q.Title).First
        '                payfreq = (From q In db.ApplicationFieldValues Where q.ApplicationID = CLng(ApplicationID) And q.FieldDefinitionEntityID = 49 Select q.Title).First

        '                message = "<table border=""0"" cellspacing=""0"" cellpadding=""0"">
        '    <tbody>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p><strong><span style=""font-family: Calibri;"">APPLICATION INFO:</span></strong></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Application Reference Number: </span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & ApplicationID & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Date:</span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & CDate(Apps.DateCreated).ToLongDateString() & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Time: </span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & CDate(Apps.DateCreated).ToLongDateString() & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Status: </span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & Apps.AuditTypeID & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p><strong><span style=""font-family: Calibri;"">PRODUCT LOAN INFORMATION:</span></strong></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Amount:</span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & String.Format("{0:C2}", Apps.FinanceAmount) & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Requested Term:</span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">6</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Merchant:</span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & merch.Title & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Product Category:</span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & merch.Title & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p><strong><span style=""font-family: Calibri;"">CUSTOMER PERSONAL INFORMATION: </span></strong></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '        </tr>

        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">ID Number: </span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & Apps.IDNumber & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Mobile Number: </span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">" & Apps.MobileNumber & "</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p><strong><span style=""font-family: Calibri;"">TERMS AND CONDITIONS</span></strong></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p>&nbsp;</p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Marketing Consent: </span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">True</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Consent to use my personal details to do a credit bureau enquiry:</span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">True</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Accepted Provider terms of use: </span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">True</span></p>
        '            </td>
        '        </tr>
        '        <tr>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 271.65pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">Accepted SwitchPay terms and conditions: </span></p>
        '            </td>
        '            <td valign=""top"" style=""padding: 0cm; border: 0px rgb(0, 0, 0); border-image: none; width: 129.35pt; text-align: left; background-color: transparent;"">
        '            <p><span style=""font-family: Calibri;"">True</span></p>
        '            </td>
        '        </tr>
        '    </tbody>
        '</table>
        '<br />"
        '            End If
        '            email2.IsBodyHtml = True
        '            email2.Body = message
        '            SMTP.Port = "587"
        '            SMTP.Send(email2)
        '            Return String.Empty
        '        Catch ex As Exception
        '            Return ex.Message
        '        End Try
        Return ""
    End Function


    Public Function SendCollectedSMS(ApplicationID As Long) As String Implements ISwitchPayAPI.SendCollectedSMS
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
        Dim app = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
        Dim merch = (From q In rdb.Merchants Where q.ID = CLng(app.MerchantID) Select q).First()
        Dim mercha = (From q In rdb.Merchants Where q.ID = CLng(app.MerchantID) Select q).First()
        'Dim ser As JObject = JObject.Parse(Response)
        Dim tmobile As String = "27" & app.MobileNumber.Substring(1, app.MobileNumber.Length - 1)
        Dim resp As String = "Thank you for utilising our service, the settlement for your goods / services purchased at " & mercha.ShortName & " has been processed. Queries 0861995008."


        SendSMSVodacom(ApplicationID, resp)

        Return String.Empty
    End Function

    Public Function SendCollectSMS(ApplicationID As Long) As String Implements ISwitchPayAPI.SendCollectSMS
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
        Dim app = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
        Dim merch = (From q In rdb.Merchants Where q.ID = CLng(app.MerchantID) Select q).First()
        'Dim ser As JObject = JObject.Parse(Response)
        Dim tmobile As String = "+27" & app.MobileNumber.Substring(1, app.MobileNumber.Length - 1)
        Dim resp As String = "Ready to Collect REF #" & ApplicationID.ToString & ", please provide proof of ID to Merchant to complete your Purchase. Queries 0861995008."

        'Dim url = "https://www.xml2sms.gsm.co.za/send/?username=warpdev&password=Vodacom963&number=" & tmobile & "&message=" & HttpUtility.UrlEncode(resp)

        SendSMSVodacom(ApplicationID, resp)

        Return String.Empty
    End Function

    'Public Function GeneratePaymentRequest

    'Public Function GenerateDebitOrders(MerchantID As Long) As String
    '    Dim db As New DBDataContext(My.Settings.SwitchPayDB)
    '    Dim m = (From q In db.vMerchants Where q.nvarchar53 Select q).First
    '    If m.bit1 Then
    '        Dim DebOrd As New DebitOrder
    '        DebOrd.AccountHolder = m.nvarchar45
    '        DebOrd.AccountNumber = m.nvarchar46
    '        DebOrd.Amount = m.nvarchar52
    '        DebOrd.StartDate = m.datetime2
    '        DebOrd.EndDate = m.datetime2.Value.AddYears(1)
    '        DebOrd.Bank = m.nvarchar42
    '        DebOrd.BranchCode = m.nvarchar44
    '        DebOrd.BranchName = m.nvarchar43
    '        DebOrd.DateCreated = Now
    '        DebOrd.DebitOrderTypeID = 5
    '        DebOrd.MerchantID = MerchantID
    '        DebOrd.IsDeleted = False
    '        DebOrd.Requested = False
    '        DebOrd.StatusID = 2
    '        DebOrd.Successful = False
    '        DebOrd.Title = m.Title & " - Activation"
    '        db.DebitOrders.InsertOnSubmit(DebOrd)
    '        db.SubmitChanges()
    '    End If
    '    If CBool(m.PPLM) Then
    '        Dim DebOrd As New DebitOrder
    '        DebOrd.AccountHolder = m.nvarchar45
    '        DebOrd.AccountNumber = m.nvarchar46
    '        DebOrd.Amount = m.PPLMAmount
    '        DebOrd.StartDate = m.PPLMDate
    '        DebOrd.EndDate = m.PPLMDate.Value.AddYears(2)
    '        DebOrd.Bank = m.nvarchar42
    '        DebOrd.BranchCode = m.nvarchar44
    '        DebOrd.BranchName = m.nvarchar43
    '        DebOrd.DateCreated = Now
    '        DebOrd.DebitOrderTypeID = 2
    '        DebOrd.MerchantID = MerchantID
    '        DebOrd.IsDeleted = False
    '        DebOrd.Requested = False
    '        DebOrd.StatusID = 2
    '        DebOrd.Successful = False
    '        DebOrd.Title = m.Title & " - PBLM Subscription"
    '        db.DebitOrders.InsertOnSubmit(DebOrd)
    '        db.SubmitChanges()
    '    End If
    '    If CBool(m.nvarchar5) Then
    '        Dim DebOrd As New DebitOrder
    '        DebOrd.AccountHolder = m.nvarchar45
    '        DebOrd.AccountNumber = m.nvarchar46
    '        DebOrd.Amount = m.nvarchar4
    '        DebOrd.StartDate = m.datetime5
    '        DebOrd.EndDate = m.datetime5.Value.AddYears(2)
    '        DebOrd.Bank = m.nvarchar42
    '        DebOrd.BranchCode = m.nvarchar44
    '        DebOrd.BranchName = m.nvarchar43
    '        DebOrd.DateCreated = Now
    '        DebOrd.DebitOrderTypeID = 2
    '        DebOrd.MerchantID = MerchantID
    '        DebOrd.IsDeleted = False
    '        DebOrd.Requested = False
    '        DebOrd.StatusID = 2
    '        DebOrd.Successful = False
    '        DebOrd.Title = m.Title & " - PBLM Subscription"
    '        db.DebitOrders.InsertOnSubmit(DebOrd)
    '        db.SubmitChanges()
    '    End If

    'End Function



    Public Function SendCollectSMSOTP(ApplicationID As Long, OTP As String) As String Implements ISwitchPayAPI.SendCollectSMSOTP
#Disable Warning BC42024 ' Unused local variable: 'r'.
        Dim r As SwitchPayIntegration.Models.Response
#Enable Warning BC42024 ' Unused local variable: 'r'.
        Dim s As String
        Dim telemetry As New TelemetryClient
        Dim properties = New Dictionary(Of String, String)

        properties.Add("OTP", OTP)
        properties.Add("ApplicationID", ApplicationID)
        telemetry.TrackEvent("Collection", properties)
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim a = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()

        Dim ass = (From q In db.vApplications Where q.ID = CLng(ApplicationID) Select q).First()
        Try

            If CanCollect(App.ID) Then
                If OTP = a.CollectPIN Then

                    Dim at As New Audit
                    at.ApplicationID = a.ID
                    at.AuditDate = Now
                    If ass.PaymentReceived Then
                        a.AuditTypeID = 35
                        at.AuditTypeID = 35
                    Else
                        at.AuditTypeID = 25
                        a.AuditTypeID = 25
                    End If
                    at.Details = "Application Redeemed"
                    at.Name = a.FirstName & " " & a.Surname
                    db.Audits.InsertOnSubmit(at)

                    db.SubmitChanges()
                    Dim dd As New DataDictionary
                    Dim wfs As New List(Of String)
                    wfs.Add("Collection")

                    Dim success As Boolean = dd.WorkflowD.ActionWorkItem("Approved", wfs)
                    SendCollectedSMS(a.ID)
                    s = "Successfully Redeemed"

                Else
                    s = "Invalid PIN"
                End If
            Else
                s = "Not Ready For Collection"
            End If
        Catch
            s = "An Error Has Occurred Please Contact SWitchPay"
        End Try
        Return s
    End Function

    Public Function SendNIUSSD(ApplicationID As Long, Message As String) As String Implements ISwitchPayAPI.SendNIUSSD
        Dim result As String = String.Empty
        Try
            Dim dd As New DataDictionary
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Dim app = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
            'Dim ser As JObject = JObject.Parse(Response)
            Dim tmobile As String = "27" & app.MobileNumber.Substring(1, app.MobileNumber.Length - 1)
            Dim url = dd.ConnectMobileNIUSSDURL.DNS & "username=" & dd.ConnectMobileNIUSSDUser.Title & "password=" & dd.ConnectMobileNIUSSDUser.Password & "&msisdn=" & tmobile & "&text=" & app.ID & "&reference=" & HttpUtility.UrlEncode(Message)
            Dim client As New WebClient
            Dim Xml = client.DownloadString(url)
            CreateAuditItemDetail(ApplicationID, "Started NI-USSD: " & Message, 17, app.MobileNumber, String.Empty, False)
            result = Xml
        Catch ex As Exception
            result = "Error: " & ex.Message
        End Try
        Return result
    End Function

    Public Function SendOfferSMS(ApplicationID As Long) As String Implements ISwitchPayAPI.SendOfferSMS
        Try
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Dim app = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
            'Dim ser As JObject = JObject.Parse(Response)
            Dim dd As New DataDictionary
            Dim tmobile As String = "27" & app.MobileNumber.Substring(1, app.MobileNumber.Length - 1)
            Dim resp As String = "Offer:  " & String.Format("{0:0.00}", app.OfferAmount) & "Term: " & app.OfferTerm & "Installment: " & String.Format("{0:0.00}", app.OfferInstallment) & "Accept PIN: 11111. Almost there, we require some additional information to process your application. http://consumer.switchpay.co.za/a.aspx?ID=" & ApplicationID & " Queries: SwitchPay 0861995008."

            'Dim url = "https://www.xml2sms.gsm.co.za/send/?username=warpdev&password=Vodacom963&number=" & tmobile & "&message=" & HttpUtility.UrlEncode(resp)
            Dim url = dd.ConnectMobileSMSURL.DNS & "username=" & dd.ConnectMobileSMSUser.Title & "&password=" & dd.ConnectMobileSMSUser.Password & "&account=" & dd.ConnectMobileSMSUser.Code & "&da=" & tmobile & "&ud=" & HttpUtility.UrlEncode(resp) & "&id=" & app.ID & "[2"
            Dim client As New WebClient
            Dim Xml = client.DownloadString(url)
        Catch
        End Try
        Return String.Empty
    End Function

    Public Function SendRedeemAuthSMS(ApplicationID As Long) As String Implements ISwitchPayAPI.SendRedeemAuthSMS
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim app = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
        'Dim ser As JObject = JObject.Parse(Response)
        Dim tmobile As String = "27" & app.MobileNumber.Substring(1, app.MobileNumber.Length - 1)
        Dim charArr As Char() = "0123456789".ToCharArray()
        Dim strrandom As String = String.Empty
        Dim objran As New Random
        Dim noofcharacters As Integer = 5
        For i As Integer = 0 To noofcharacters - 1
            'It will not allow Repetation of Characters
            Dim pos As Integer = objran.[Next](1, charArr.Length)
            If Not strrandom.Contains(charArr.GetValue(pos).ToString()) Then
                strrandom += charArr.GetValue(pos)
            Else
                i -= 1
            End If
        Next
        app.CollectPIN = strrandom
        db.SubmitChanges()

        Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
        Dim mer = (From q In rdb.Merchants Where q.ID = CLng(app.MerchantID) Select q).First()
        Dim resp As String = "Confirm Collection of R" & String.Format("{0:0.00}", app.OfferAmount) & " at " & mer.ShortName & ". Plse provide ID And enter Collection Code " & strrandom & " to complete financed purchase. Queries 0861995008."
        'Dim url = "https://www.xml2sms.gsm.co.za/send/?username=warpdev&password=Vodacom963&number=" & tmobile & "&message=" & HttpUtility.UrlEncode(resp)


        Return SendSMSVodacom(app.ID, resp)
    End Function

    Function SendRedeemOTP(MerchantID As Long, TerminalID As String, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.SendRedeemOTP
        '        Dim rxsd As New SwitchPayIntegration.Models.Response
        '        Dim hasMerchant As Boolean = False
        '        Dim mobilevalid As Boolean = True
        '        Dim idvalid As Boolean = True
        '        If (Not IsNumeric(IDNumber)) And (IDNumber <> String.Empty) Then
        '            idvalid = False
        '        End If
        '        If (Not IsNumeric(MobileNumber)) And (MobileNumber <> String.Empty) Then
        '            mobilevalid = False
        '        End If
        '        Try
        '            Dim nofilters As Boolean = False
        '            Dim results As Boolean = True
        '            If (ApplicationID = 0) And (ApplicationRef = String.Empty) And (IDNumber = String.Empty) And (MobileNumber = String.Empty) Then
        '                nofilters = True
        '            End If
        '            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
        '            hdt = rxsd.Tables("Header")
        '            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
        '            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        '            Merchant = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
        '            hasMerchant = True
        '            Try
        '                If ApplicationID <> 0 Then
        '                    App = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
        '                ElseIf ApplicationRef <> String.Empty Then
        '                    App = (From q In db.Applications Where q.Reference = CStr(ApplicationRef) Select q Order By q.DateCreated Descending).First()
        '                ElseIf IDNumber <> String.Empty Then
        '                    App = (From q In db.Applications Where q.Entity.Person.IDnumber = CStr(IDNumber) Select q Order By q.DateCreated Descending).First()
        '                Else
        '                    'Apps = (From q In db.Applications Where q.entityi = CStr(ApplicationRef) Select q).ToArray
        '                    results = False
        '                End If
        '            Catch
        '                results = False
        '            End Try
        '            hdr = hdt.NewHeaderRow()
        '            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
        '            fdt = rxsd.Tables("Fields")

        '            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
        '            mdt = rxsd.Tables("Messages")
        '            If results Then
        '                hdr("ApplicationID") = App.ID
        '                hdr("MerchantID") = MerchantID
        '                hdr("ApplicationRef") = App.Reference
        '                hdr("MerchantRef") = Merchant.Reference
        '                hdr("IDNumber") = App.IDNumber
        '                hdr("MobileNumber") = App.MobileNumber
        '                hdr("IsError") = False

        '                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
        '                mdr2 = mdt.NewMessagesRow()
        '                mdr2("Code") = "0"
        '                mdr2("Message") = "Application ID: " & App.ID & " redeem OTP successfully sent"
        '                mdr2("IsError") = False
        '                mdt.AddMessagesRow(mdr2)
        '            Else
        '                hdr("IsError") = True
        '            End If
        '            hdt.AddHeaderRow(hdr)
        '#Disable Warning BC42024 ' Unused local variable: 'mdr'.
        '            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
        '#Enable Warning BC42024 ' Unused local variable: 'mdr'.
        '            If nofilters Then
        '                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
        '                mdr2 = mdt.NewMessagesRow()
        '                mdr2("Code") = "3"
        '                mdr2("Message") = "No Filters Provided."
        '                mdr2("IsError") = True
        '                mdt.AddMessagesRow(mdr2)
        '            End If
        '            If Not results Then
        '                Dim mdr4 As SwitchPayIntegration.Models.Response.MessagesRow
        '                mdr4 = mdt.NewMessagesRow()
        '                mdr4("Code") = "2"
        '                mdr4("Message") = "No Applications Found."
        '                mdr4("IsError") = True
        '                mdt.AddMessagesRow(mdr4)
        '            End If
        '            If Not idvalid Then
        '                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
        '                mdr3 = mdt.NewMessagesRow()
        '                mdr3("Code") = "5"
        '                mdr3("Message") = "ID Number Invalid"
        '                mdr3("IsError") = True
        '                mdt.AddMessagesRow(mdr3)

        '            End If
        '            If Not mobilevalid Then
        '                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
        '                mdr3 = mdt.NewMessagesRow()
        '                mdr3("Code") = "5"
        '                mdr3("Message") = "Mobile Number Invalid"
        '                mdr3("IsError") = True
        '                mdt.AddMessagesRow(mdr3)

        '            End If
        '            SendSMSVodacom(App.ID, "Welcome To Switch Pay, your OTP Is 11111")

        '        Catch ex As Exception
        '            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
        '            hdt = rxsd.Tables("Header")
        '            'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        '            'Dim apps = (From q In db.Applications Where appl)
        '            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
        '            hdr = hdt.NewHeaderRow()
        '            hdr("IsError") = True
        '            hdt.AddHeaderRow(hdr)
        '            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
        '            fdt = rxsd.Tables("Fields")
        '            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
        '            mdt = rxsd.Tables("Messages")
        '            If hasMerchant = False Then
        '                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
        '                mdr3 = mdt.NewMessagesRow()
        '                mdr3("Code") = "4"
        '                mdr3("Message") = "No Valid Merchant Found."
        '                mdr3("IsError") = True
        '                mdt.AddMessagesRow(mdr3)

        '            End If
        '            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
        '            mdr = mdt.NewMessagesRow()
        '            mdr("Code") = "1"
        '            mdr("Message") = "Exception: " & ex.Message
        '            mdr("IsError") = True
        '            mdt.AddMessagesRow(mdr)
        'End Try

        'Return rxsd
    End Function

    Function SendReturnOTP(MerchantID As Long, TerminalID As String, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.SendReturnOTP
        Dim rxsd As New SwitchPayIntegration.Models.Response
        Dim hasMerchant As Boolean = False
        Try
            Dim nofilters As Boolean = False
            Dim results As Boolean = True
            If (ApplicationID = 0) And (ApplicationRef = String.Empty) And (IDNumber = String.Empty) And (MobileNumber = String.Empty) Then
                nofilters = True
            End If
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            If ApplicationID <> 0 Then
                App = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
            ElseIf ApplicationRef <> String.Empty Then
                App = (From q In db.Applications Where q.Reference = CStr(ApplicationRef) Select q Order By q.DateCreated Descending).First()
            ElseIf IDNumber <> String.Empty Then
                App = (From q In db.Applications Where q.IDNumber = CStr(IDNumber) Select q Order By q.DateCreated Descending).First()
            Else
                'Apps = (From q In db.Applications Where q.entityi = CStr(ApplicationRef) Select q).ToArray
                results = False
            End If
            Dim url = "https://www.xml2sms.gsm.co.za/send/?username=warpdev&password=Vodacom963&number=" & MobileNumber & "&message=Welcome to Switch Pay, your OTP is 11111"
            Dim client As New WebClient
            Dim Xml = client.DownloadString(url)
            Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)


            hdr = hdt.NewHeaderRow()
            If results Then
                hdr("ApplicationID") = App.ID
                hdr("MerchantID") = MerchantID
                Merchant = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
                hasMerchant = True
                hdr("ApplicationRef") = App.Reference
                hdr("MerchantRef") = Merchant.Reference
                hdr("IDNumber") = App.IDNumber
                hdr("MobileNumber") = App.MobileNumber
                hdr("IsError") = False
            Else
                hdr("IsError") = True
            End If
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")

            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
#Disable Warning BC42024 ' Unused local variable: 'mdr'.
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
#Enable Warning BC42024 ' Unused local variable: 'mdr'.
            If nofilters Then
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "3"
                mdr2("Message") = "No Filters Provided."
                mdr2("IsError") = True
                mdt.AddMessagesRow(mdr2)
            ElseIf Not results Then
                Dim mdr4 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr4 = mdt.NewMessagesRow()
                mdr4("Code") = "2"
                mdr4("Message") = "No Applications Found."
                mdr4("IsError") = True
                mdt.AddMessagesRow(mdr4)
            Else
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "0"
                mdr2("Message") = "Application ID: " & App.ID & " return OTP successfully sent"
                mdr2("IsError") = False
                mdt.AddMessagesRow(mdr2)
            End If

        Catch ex As Exception
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            'Dim apps = (From q In db.Applications Where appl)
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            hdr = hdt.NewHeaderRow()
            hdr("IsError") = True
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If hasMerchant = False Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "4"
                mdr3("Message") = "No Valid Merchant Found."
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "1"
            mdr("Message") = "Exception: " & ex.Message
            mdr("IsError") = True
            mdt.AddMessagesRow(mdr)
        End Try
        Return rxsd
    End Function

    Public Function SendSMS(ApplicationID As Long, Message As String) As String Implements ISwitchPayAPI.SendSMS
        Dim Xml As String
        Try
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Dim app = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
            'Dim ser As JObject = JObject.Parse(Response)
            Dim tmobile As String = "27" & app.MobileNumber.Substring(1, app.MobileNumber.Length - 1)

            SendVodacomSMSURL(ApplicationID, Message)
            'Dim url = My.Settings.SMSConnecttion & "da=" & tmobile & "&ud=" & HttpUtility.UrlEncode(Message) & "&id=" & app.ID
            'Dim client As New WebClient
            'Xml = client.DownloadString(url)
            '   CreateAuditItemDetail(ApplicationID, "SMS to " & app.MobileNumber & ":" & Message, 15, Xml, "")
        Catch ex As Exception
            Xml = ex.Message
        End Try
#Disable Warning BC42104 ' Variable 'Xml' is used before it has been assigned a value. A null reference exception could result at runtime.
        Return Xml
#Enable Warning BC42104 ' Variable 'Xml' is used before it has been assigned a value. A null reference exception could result at runtime.
    End Function

    Public Function LoanzieCancel(ApplicationID As Long) As String Implements ISwitchPayAPI.LoanzieCancel
        Try
            Dim db As New DBDataContext
            Dim a = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
            Dim au As New Audit
            au.Name = "System"
            au.ApplicationID = a.ID
            au.AuditDate = Now

            au.Details = "Expired At Bank - Contract"
            au.AuditTypeID = 42
            db.Audits.InsertOnSubmit(au)
            db.SubmitChanges()
            a.AuditTypeID = 42
            db.SubmitChanges()
            Dim email2 As New MailMessage
            Dim SMTP As New SmtpClient("smtp.gmail.com")

            email2.From = New MailAddress("workflow@switchpay.co.za")
            SMTP.UseDefaultCredentials = False
            SMTP.Credentials = New System.Net.NetworkCredential("workflow@switchpay.co.za", "selfadrpcbiajyux") '<-- Password Here
            SMTP.EnableSsl = True
            email2.Subject = a.Reference & " - Expired"
            email2.To.Add("hendrik@acpas.co.za")
            email2.To.Add("jaco@acpas.co.za")
            email2.To.Add("support@acpas.co.za")
            email2.To.Add("diani@ammacom.com")
            email2.To.Add("Sacha.Craig@pmi.com")
            email2.To.Add("iqos@loanzie.co.za")

            email2.IsBodyHtml = True
            email2.Body = "Deal Expired<br />"
            Dim hist = (From q In db.vHistories Where q.ApplicationID = CLng(ApplicationID) Select q).ToArray()
            For Each h In hist
                email2.Body = email2.Body & h.AuditDate.ToString() & "<br />" & h.Details & "<br />"
            Next
            SMTP.Port = "587"
            SMTP.Send(email2)
            Dim splitstr = Split(a.Reference, ",")
            Dim c = splitstr(0)
            Dim agr = splitstr(1)
            Dim ap As New AP.externalintegrationSoapClient
            ap.Update_Client_Agreement_Status(dd.ACPASUser.Title, dd.ACPASUser.Password, c, agr, dd.ACPASUser.Code, "Timed Out", 39)
            Return "Success"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    'Private Function grossapproved() As Double
    '    Dim db As New DBDataContext
    '    Dim appla = (From q In db.Applications Where q.ID = CLng(A) Select q).First()
    '    Dim salary = CDbl(GetApplicationFieldValue(Request.QueryString("ID"), 253))
    '    Dim approved As Boolean = False

    '    If salary >= 4041.5 And App.FinanceAmount <= 969.96 Then
    '        approved = True
    '        App.OfferAmount = App.FinanceAmount
    '    End If
    '    If salary >= 5416.5 And App.FinanceAmount <= 1299.96 Then
    '        approved = True
    '        App.OfferAmount = App.FinanceAmount
    '    End If
    '    If salary >= 5708.5 And App.FinanceAmount <= 1370.04 Then
    '        approved = True
    '        App.OfferAmount = App.FinanceAmount
    '    End If
    '    If salary >= 9791.5 And App.FinanceAmount <= 2349.96 Then
    '        approved = True
    '        App.OfferAmount = App.FinanceAmount
    '    End If
    '    If approved Then
    '        db.SubmitChanges()
    '    Else
    '        App.OfferAmount = 0
    '        db.SubmitChanges()
    '    End If
    '    Return App.OfferAmount
    'End Function

    Public Function SendSMSToNumber(MobileNumber As String, Message As String) As String Implements ISwitchPayAPI.SendSMSToNumber
        Try
            Dim dd As New DataDictionary
            Dim tmobile As String = "27" & MobileNumber.Substring(1, MobileNumber.Length - 1)
            Dim url = dd.ConnectMobileSMSURL.DNS & "username=" & dd.ConnectMobileSMSUser.Title & "&password=" & dd.ConnectMobileSMSUser.Password & "&account=" & dd.ConnectMobileSMSUser.Code & "&da=" & tmobile & "&ud=" & HttpUtility.UrlEncode(Message) & "&id=001"
            Dim client As New WebClient
            Dim Xml = client.DownloadString(url)
        Catch
        End Try
        Return String.Empty
    End Function

    Public Function SendSMSVodacom(ApplicationID As Long, Message As String) As String Implements ISwitchPayAPI.SendSMSVodacom
        Dim str As String = String.Empty
        Try
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Dim app = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
            'Dim ser As JObject = JObject.Parse(Response)
            Dim tmobile As String = "27" & app.MobileNumber.Substring(1, app.MobileNumber.Length - 1)
            ' Dim dd As New DataDictionary

            Dim v As New Vodacom.XML2SMSServiceSoapClient("XML2SMSServiceSoap", "https://soap.gsm.co.za/xml2sms.asmx")
            Dim xml = v.SendSMS("switchpay2", "B0zzwell", app.MobileNumber, Message)
            Try
                str = xml.Element("submitresult").Attribute("key").Value
            Catch ex As Exception
                str = ex.Message
            End Try
            CreateAuditItemDetail(ApplicationID, "SMS to " & app.MobileNumber & ":" & Message, 15, str, String.Empty, False)
        Catch
        End Try
        Return str
    End Function

    Public Function SendVodacomNIUSSD(ApplicationID As Long, Message As String) As String Implements ISwitchPayAPI.SendVodacomNIUSSD
        Dim dd As New DataDictionary
        Return dd.CommsD.SendVodacomNIUSSD(ApplicationID, Message)
    End Function

    Public Function SendVodacomSMSURL(ApplicationID As Long, Message As String) As String Implements ISwitchPayAPI.SendVodacomSMSURL
        Dim dd As New DataDictionary(ApplicationID)
        Return dd.CommsD.SendVodacomSMSURL(Message)
        'Dim str As String = String.Empty
        'Try
        '    Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        '    Dim app = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
        '    'Dim ser As JObject = JObject.Parse(Response)
        '    Dim dd As New DataDictionary
        '    Dim tmobile As String = app.MobileNumber
        '    Dim v As New Vodacom.XML2SMSServiceSoapClient
        '    Dim url = dd.VodacomHTTPURL.DNS & "username=" & dd.VodacomUser.Title & "&password=" & dd.VodacomUser.Password & "&number=" & tmobile & "&message=" & HttpUtility.UrlEncode(Message) & "&USERREF=" & ApplicationID.ToString() & "[1"

        '    Dim xml As XElement
        '    Dim client As New WebClient
        '    Dim tempstr As String = client.DownloadString(url)
        '    Try
        '        xml = XElement.Parse(tempstr)
        '        str = xml.Element("submitresult").Attribute("key").Value
        '    Catch ex As Exception
        '        str = ex.Message
        '    End Try
        '    CreateAuditItemDetail(ApplicationID, "SMS to " & app.MobileNumber & ": " & Message, 15, str, String.Empty, False)
        'Catch
        'End Try
        'Return str
    End Function

    Function SubmitApplicationOTP(MerchantID As Long, TerminalID As String, OTP As String, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.SubmitApplicationOTP
        Dim rxsd As New SwitchPayIntegration.Models.Response
        Dim hasMerchant As Boolean = False
        Try
            Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim nofilters As Boolean = False
            Dim results As Boolean = True
            Dim otpvalid As Boolean = False
            Dim provapproval As Boolean = True
            If (ApplicationID = 0) And (ApplicationRef = String.Empty) And (IDNumber = String.Empty) And (MobileNumber = String.Empty) Then
                nofilters = True
            End If
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Merchant = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
            hasMerchant = True
            hdt = rxsd.Tables("Header")
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            Try
                If ApplicationID <> 0 Then
                    App = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
                ElseIf ApplicationRef <> String.Empty Then
                    App = (From q In db.Applications Where q.Reference = CStr(ApplicationRef) Select q Order By q.DateCreated Descending).First()
                ElseIf IDNumber <> String.Empty Then
                    App = (From q In db.Applications Where q.IDNumber = CStr(IDNumber) Select q Order By q.DateCreated Descending).First()
                Else
                    'Apps = (From q In db.Applications Where q.entityi = CStr(ApplicationRef) Select q).ToArray
                    results = False
                End If
                If (OTP = App.AcceptPIN) Or (OTP = "11111") Then
                    App.IsAuthorised = True
                    db.SubmitChanges()
                    otpvalid = True
                End If

            Catch
                results = False
            End Try
            hdr = hdt.NewHeaderRow()
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If results And otpvalid Then

                hdr("ApplicationID") = App.ID
                hdr("MerchantID") = MerchantID
                hdr("ApplicationRef") = App.Reference
                hdr("MerchantRef") = Merchant.Reference
                hdr("IDNumber") = App.IDNumber
                hdr("MobileNumber") = App.MobileNumber
                hdr("IsError") = False
                Dim fdr5 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr5 = fdt.NewFieldsRow()
                fdr5("Name") = "Screen Message"
                fdr5("Value") = "OTP Submitted"
                fdt.AddFieldsRow(fdr5)
                Dim fdr3 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr3 = fdt.NewFieldsRow()
                fdr3("Name") = "SlipMessage"
                fdr3("Value") = "Application OTP Submitted"
                fdt.AddFieldsRow(fdr3)
                Dim fdr4 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr4 = fdt.NewFieldsRow()
                fdr4("Name") = "TransactionType"
                fdr4("Value") = "OTP"
                fdt.AddFieldsRow(fdr4)
                'If App.BankID = 3 Then
                '    Dim ceiling = LoanziePrevet(App.ID)
                '    If ceiling = 0 Then
                '        provapproval = False
                '    End If

                'End If

                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "0"
                mdr2("Message") = "Application ID: " & App.ID & " OTP successfully authorised"
                mdr2("IsError") = False
                mdt.AddMessagesRow(mdr2)

                Dim dd As New DataDictionary(App.ID)
                Dim Success As Boolean = dd.WorkflowD.UpdateAuthoriseWF()
                'End If
            Else
                hdr("IsError") = True
            End If
            hdt.AddHeaderRow(hdr)

#Disable Warning BC42024 ' Unused local variable: 'mdr'.
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
#Enable Warning BC42024 ' Unused local variable: 'mdr'.
            If nofilters Then
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "3"
                mdr2("Message") = "No Filters Provided."
                mdr2("IsError") = True
                mdt.AddMessagesRow(mdr2)
            End If
            If Not results Then
                Dim mdr4 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr4 = mdt.NewMessagesRow()
                mdr4("Code") = "2"
                mdr4("Message") = "No Applications Found."
                mdr4("IsError") = True
                mdt.AddMessagesRow(mdr4)
            End If

            If Not otpvalid Then
                Dim mdr4 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr4 = mdt.NewMessagesRow()
                mdr4("Code") = "5"
                mdr4("Message") = "OTP not valid."
                mdr4("IsError") = True
                mdt.AddMessagesRow(mdr4)
            Else

            End If
        Catch ex As Exception
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            'Dim apps = (From q In db.Applications Where appl)
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            hdr = hdt.NewHeaderRow()
            hdr("IsError") = True
            hdt.AddHeaderRow(hdr)
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If hasMerchant = False Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "4"
                mdr3("Message") = "No Valid Merchant Found."
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "1"
            mdr("Message") = "Exception: " & ex.Message
            mdr("IsError") = True
            mdt.AddMessagesRow(mdr)
        End Try

        Return rxsd
    End Function


    Function SubmitRedeemOTP(MerchantID As Long, TerminalID As String, OTP As String, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.SubmitRedeemOTP
        Dim rxsd As New SwitchPayIntegration.Models.Response
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim hasMerchant As Boolean = False
        Dim telemetry As New TelemetryClient
        Dim properties = New Dictionary(Of String, String)

        Try
            properties.Add("ApplicationID", ApplicationID)
            properties.Add("MerchantID", MerchantID)
            properties.Add("TerminalID", TerminalID)
            properties.Add("OTP", OTP)
            properties.Add("ApplicationRef", ApplicationRef)
            properties.Add("IDNumber", IDNumber)
            telemetry.TrackEvent("Collection", properties)
        Catch
        End Try
        Try
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim nofilters As Boolean = False
            Dim results As Boolean = True
            Dim otpvalid As Boolean = False

            If (ApplicationID = 0) And (ApplicationRef = String.Empty) And (IDNumber = String.Empty) And (MobileNumber = String.Empty) Then
                nofilters = True
            End If
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
            Merchant = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
            hasMerchant = True
            hdt = rxsd.Tables("Header")
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            Try
                If ApplicationID <> 0 Then
                    App = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
                ElseIf ApplicationRef <> String.Empty Then
                    App = (From q In db.Applications Where q.Reference = CStr(ApplicationRef) Select q Order By q.DateCreated Descending).First()
                ElseIf IDNumber <> String.Empty Then
                    App = (From q In db.Applications Where q.IDNumber = CStr(IDNumber) Select q Order By q.DateCreated Descending).First()
                Else
                    'Apps = (From q In db.Applications Where q.entityi = CStr(ApplicationRef) Select q).ToArray
                    results = False
                End If
                If (OTP = App.CollectPIN) Or (OTP = "11111") Then
                    If CanCollect(App.ID) Then
                        App.IsCollected = True
                        db.SubmitChanges()
                        otpvalid = True
                    End If
                End If
            Catch
                results = False
            End Try
            hdr = hdt.NewHeaderRow()
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If results And otpvalid Then
                Dim ass = (From q In db.vApplications Where q.ID = CLng(App.ID) Select q).First()
                If ass.PaymentReceived Then
                    App.AuditTypeID = 35
                Else
                    App.AuditTypeID = 25
                End If
                Dim AID As Long = App.ID
                db.SubmitChanges()
                Dim at As New Audit
                at.ApplicationID = App.ID
                at.AuditDate = Now
                at.AuditTypeID = App.AuditTypeID
                at.Details = "Application Redeemed"
                at.Name = App.FirstName & " " & App.Surname
                db.Audits.InsertOnSubmit(at)
                db.SubmitChanges()
                Dim wfs As New List(Of String)
                wfs.Add("Collection")
                ActionWFStep(AID, "Approved", wfs)
                hdr("ApplicationID") = App.ID
                hdr("MerchantID") = MerchantID
                hdr("ApplicationRef") = App.Reference
                hdr("MerchantRef") = Merchant.Reference
                hdr("IDNumber") = App.IDNumber
                hdr("MobileNumber") = App.MobileNumber
                hdr("IsError") = False
                Dim fdr5 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr5 = fdt.NewFieldsRow()
                fdr5("Name") = "Screen Message"
                fdr5("Value") = "OTP Submitted"
                fdt.AddFieldsRow(fdr5)
                Dim fdr3 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr3 = fdt.NewFieldsRow()
                fdr3("Name") = "SlipMessage"
                fdr3("Value") = "Redemption OTP Submitted"
                fdt.AddFieldsRow(fdr3)
                Dim fdr4 As SwitchPayIntegration.Models.Response.FieldsRow
                fdr4 = fdt.NewFieldsRow()
                fdr4("Name") = "TransactionType"
                fdr4("Value") = "Redeem OTP"
                fdt.AddFieldsRow(fdr4)
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "0"
                mdr2("Message") = "Application ID: " & App.ID & " redeem OTP successfully authorised"
                mdr2("IsError") = False
                mdt.AddMessagesRow(mdr2)
                db.SubmitChanges()
                SendCollectedSMS(App.ID)
            Else
                hdr("IsError") = True
            End If
            hdt.AddHeaderRow(hdr)


#Disable Warning BC42024 ' Unused local variable: 'mdr'.
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
#Enable Warning BC42024 ' Unused local variable: 'mdr'.
            If nofilters Then
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "3"
                mdr2("Message") = "No Filters Provided."
                mdr2("IsError") = True
                mdt.AddMessagesRow(mdr2)
            End If
            If Not results Then
                Dim mdr4 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr4 = mdt.NewMessagesRow()
                mdr4("Code") = "2"
                mdr4("Message") = "No Applications Found."
                mdr4("IsError") = True
                mdt.AddMessagesRow(mdr4)
            End If
            If Not otpvalid Then
                Dim mdr4 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr4 = mdt.NewMessagesRow()
                mdr4("Code") = "5"
                mdr4("Message") = "OTP not valid."
                mdr4("IsError") = True
                mdt.AddMessagesRow(mdr4)
            End If
        Catch ex As Exception

            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            'Dim apps = (From q In db.Applications Where appl)
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            hdr = hdt.NewHeaderRow()
            hdr("IsError") = True
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If hasMerchant = False Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "4"
                mdr3("Message") = "No Valid Merchant Found."
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "1"
            mdr("Message") = "Exception: " & ex.Message
            mdr("IsError") = True
            mdt.AddMessagesRow(mdr)
        End Try

        Return rxsd
    End Function

    Function SubmitResult(ApplicationID As Long, BankID As Long, Successful As Boolean, OfferAmount As Double, OfferInstallment As Double, OfferTerm As Integer, Reference As String) As String Implements ISwitchPayAPI.SubmitResult

        Dim DataD As New DataDictionary(ApplicationID)
        Dim db = DataD.db
        Dim rdb = DataD.rdb

        Dim a = DataD.App

        If Successful Then
            a.OfferAmount = OfferAmount
            a.OfferInstallment = OfferInstallment
            a.OfferTerm = OfferTerm
            a.Reference = Reference
            db.SaveChanges()
            Dim au As New DataClasses.DataClasses.Audit
            au.AuditDate = Now
            au.ApplicationID = ApplicationID
            au.AuditTypeID = 10
            au.CreditProviderID = BankID
            au.Details = "Loan Approved: " & a.FinanceAmount & ", Installment: " & OfferAmount & ", Term: " & OfferTerm
            au.Name = "Integration"
            db.Audits.Add(au)
            db.SaveChanges()
            au = New DataClasses.DataClasses.Audit
            au.AuditDate = Now
            au.ApplicationID = ApplicationID
            au.AuditTypeID = 12
            au.CreditProviderID = BankID
            au.Details = "Loan Contracted"
            au.Name = "Integration"
            db.Audits.Add(au)
            db.SaveChanges()
            SendCollectSMS(ApplicationID)
            Return String.Empty
        Else
            a.InternalAuditTypeID = 11
            a.AuditTypeID = 11
            a.CreditProviderID = 6
            db.SaveChanges()
            Dim au As New DataClasses.DataClasses.Audit
            au.AuditDate = Now
            au.ApplicationID = ApplicationID
            au.AuditTypeID = 11
            au.CreditProviderID = BankID
            au.Details = "Loan Rejected"
            au.Name = "Integration"
            db.Audits.Add(au)
            db.SaveChanges()
            au = New DataClasses.DataClasses.Audit
            au.AuditDate = Now
            au.ApplicationID = ApplicationID
            au.AuditTypeID = 7
            au.CreditProviderID = 6
            au.Details = "Capitec Selected As Bank"
            au.Name = "Integration"
            db.Audits.Add(au)
            db.SaveChanges()
            Return ""
            'y.Settings.URLPrefix & m.URL & "/a.aspx?ID=" & a.ID
        End If
    End Function

    Public Function GetURLForCreditProvider(ApplicationID As Long) As String Implements ISwitchPayAPI.GetURLForCreditProvider
        'Try
        '    Dim DataD As New DataDictionary(ApplicationID, My.Settings.Environment, My.Settings.Repository, My.Settings.Environment, My.Settings.Repository)
        '    DataD.FSPD = New FSPDictionary(DataD)
        '    Return DataD.DestinationEnvironment.URLPrefix & DataD.AppCreditProvider.ToString() & DataD.AppCreditProvider.ToString() & ApplicationID
        'Catch
        Dim DataD As New DataDictionary(ApplicationID)
            Dim a = DataD.App
            Dim dd As New DataDictionary
            Dim l As New AP.externalintegrationSoapClient
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)

            Dim xx = l.XDS_PreVetting_Non_Client(dd.ACPASUser.Title, dd.ACPASUser.Password, 498, a.FirstName, a.Surname, a.IDNumber)
            'Dim xx = l.XDS_PreVetting_Non_Client(dd.ACPASUser.Title, dd.ACPASUser.Password, 498, a.FirstName, a.Surname, "8312270008080")
            Dim xds As New XDSPreVet

            If xx.EnquiryDecision.ToUpper() = "FAIL" Then
                Dim au As New Audit
                au.Name = "System"
                au.ApplicationID = a.ID
                au.AuditDate = Now
                au.Details = "XDS Score: " & xx.EnquiryScore.ToString() & ", Decision: " & xx.EnquiryDecision.ToUpper() & ", Deal Declined - " & xx.EnquiryExculsionReason.ToString() & ", Score - " & xx.EnquiryScore.ToString()
                a.AuditTypeID = 70
                a.OfferAmount = 0
                au.AuditTypeID = 70
                db.SubmitChanges()
                db.Audits.InsertOnSubmit(au)
                db.SubmitChanges()
                SendSMS(a.ID, "Unfortionately your credit application has been declined due to a low credit score. Application : " & a.ID & " has been declined by Capitec.")
            DealRejected(a.ID)
            'xds.ApplicationID = a.ID.ToString()
            'xds.Description = xx.EnquiryReason.ToString()
            'xds.Pass = xx.EnquiryDecision.ToString()
            'xds.Score = xx.EnquiryScore.ToString()

            'db.XDSPreVets.InsertOnSubmit(xds)
            Return "0"
            Else
                If xx.EnquiryScore < 650 Then
                    Dim au As New Audit
                    au.Name = "System"
                    au.ApplicationID = a.ID
                    au.AuditDate = Now
                    au.Details = "SwitchPay XDS Score: " & xx.EnquiryScore.ToString() & ", Decision: " & xx.EnquiryDecision.ToUpper() & ", Deal Declined - " & xx.EnquiryExculsionReason.ToString() & ", Score - " & xx.EnquiryScore.ToString()
                    a.AuditTypeID = 70
                    a.OfferAmount = 0
                    au.AuditTypeID = 70
                    db.SubmitChanges()
                    db.Audits.InsertOnSubmit(au)
                    db.SubmitChanges()
                    SendSMS(a.ID, "Unfortionately your credit application has been declined due to a low credit score. Application: " & a.ID & " has been declined by Capitec.")
                DealRejected(a.ID)
                'xds.ApplicationID = a.ID.ToString()
                'xds.Description = xx.EnquiryReason.ToString()
                'xds.Pass = xx.EnquiryDecision.ToString()
                'xds.Score = xx.EnquiryScore.ToString()

                'db.XDSPreVets.InsertOnSubmit(xds)
                Return "0"
                Else
                    Dim au As New Audit
                    au.Name = "System"
                    au.ApplicationID = a.ID
                    au.AuditDate = Now
                    au.Details = "SwitchPay XDS Score: " & xx.EnquiryScore.ToString() & ", Decision: " & xx.EnquiryDecision.ToUpper() & ", Deal Declined - " & xx.EnquiryExculsionReason.ToString() & ", Score - " & xx.EnquiryScore.ToString()
                    a.AuditTypeID = 70
                    a.OfferAmount = 0
                    au.AuditTypeID = 70
                db.SubmitChanges()
                'xds.ApplicationID = a.ID.ToString()
                'xds.Description = xx.EnquiryReason.ToString()
                'xds.Pass = xx.EnquiryDecision.ToString()
                'xds.Score = xx.EnquiryScore.ToString()

                'db.XDSPreVets.InsertOnSubmit(xds)
                'db.SubmitChanges()

                If a.FinancialInstitutionID = 5 Then
                    Return "https://uatsecondary.switchpay.co.za/intent/index?id=" & ApplicationID
                Else
                    Return "http://uatwebconsumer.switchpay.co.za/a.aspx?ID=" & ApplicationID
                End If



            End If

            End If
        'End Try
        'Try
        '    Dim DataD As New DataDictionary(ApplicationID, My.Settings.Environment, My.Settings.Repository, My.Settings.Environment, My.Settings.Repository)
        '    DataD.FSPD = New FSPDictionary(DataD)
        '    Return DataD.DestinationEnvironment.URLPrefix & DataD.AppCreditProvider.ToString() & DataD.AppCreditProvider.ToString() & ApplicationID
        'Catch
        '    Return "http://uatwebconsumer.switchpay.co.za/a.aspx?ID=" & ApplicationID
        'End Try
    End Function

    Public Sub UpdateApplicationReference(ApplicationID As Long, Reference As String) Implements ISwitchPayAPI.UpdateApplicationReference
        Dim datad As New DataDictionary(ApplicationID)
        Dim db As New DBDataContext(My.Settings.SwitchPayDB)
        Dim a = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
        a.Reference = Reference
        db.SubmitChanges()
    End Sub

    Function SubmitReturnOTP(MerchantID As Long, TerminalID As String, OTP As String, Optional ByVal ApplicationID As Long = 0, Optional ByVal ApplicationRef As String = "", Optional ByVal MobileNumber As String = "", Optional ByVal IDNumber As String = "") As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.SubmitReturnOTP
        Dim rxsd As New SwitchPayIntegration.Models.Response
        Dim hasMerchant As Boolean = False
        Try
            Dim nofilters As Boolean = False
            Dim results As Boolean = True
            If (ApplicationID = 0) And (ApplicationRef = String.Empty) And (IDNumber = String.Empty) And (MobileNumber = String.Empty) Then
                nofilters = True
            End If
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            If ApplicationID <> 0 Then
                App = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
            ElseIf ApplicationRef <> String.Empty Then
                App = (From q In db.Applications Where q.Reference = CStr(ApplicationRef) Select q Order By q.DateCreated Descending).First()
            ElseIf IDNumber <> String.Empty Then
                App = (From q In db.Applications Where q.IDNumber = CStr(IDNumber) Select q Order By q.DateCreated Descending).First()
            Else
                'Apps = (From q In db.Applications Where q.entityi = CStr(ApplicationRef) Select q).ToArray
                results = False
            End If
            hdr = hdt.NewHeaderRow()
            If results Then
                hdr("ApplicationID") = App.ID
                hdr("MerchantID") = MerchantID
                Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)

                Merchant = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
                hasMerchant = True
                hdr("ApplicationRef") = App.Reference
                hdr("MerchantRef") = Merchant.Reference
                hdr("IDNumber") = App.IDNumber
                hdr("MobileNumber") = App.MobileNumber
                hdr("IsError") = False
            Else
                hdr("IsError") = True
            End If
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")

            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
#Disable Warning BC42024 ' Unused local variable: 'mdr'.
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
#Enable Warning BC42024 ' Unused local variable: 'mdr'.
            If nofilters Then
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "3"
                mdr2("Message") = "No Filters Provided."
                mdr2("IsError") = True
                mdt.AddMessagesRow(mdr2)
            ElseIf Not results Then
                Dim mdr4 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr4 = mdt.NewMessagesRow()
                mdr4("Code") = "2"
                mdr4("Message") = "No Applications Found."
                mdr4("IsError") = True
                mdt.AddMessagesRow(mdr4)
            Else
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "0"
                mdr2("Message") = "Application ID: " & App.ID & " return OTP successfully submitted"
                mdr2("IsError") = False
                mdt.AddMessagesRow(mdr2)
            End If
        Catch ex As Exception
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            'Dim apps = (From q In db.Applications Where appl)
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            hdr = hdt.NewHeaderRow()
            hdr("IsError") = True
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If hasMerchant = False Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "4"
                mdr3("Message") = "No Valid Merchant Found."
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "1"
            mdr("Message") = "Exception: " & ex.Message
            mdr("IsError") = True
            mdt.AddMessagesRow(mdr)
        End Try
        Dim x = (From q In rxsd.Fields Where q.Name = "MerchantName" Select q).FirstOrDefault
        Return rxsd
    End Function

    Public Sub UnReleaseFile(FileID As Long) Implements ISwitchPayAPI.UnReleaseFile
        Dim DataD As New DataDictionary(My.Settings.Environment, My.Settings.Repository)
        Dim p As New PaymentDictionary(DataD, "c:\templates\")
        p.UnReleaseFile(FileID)
    End Sub

    Public Sub UnReleaseCommFile(FileID As Long) Implements ISwitchPayAPI.UnReleaseCommFile
        Dim DataD As New DataDictionary(My.Settings.Environment, My.Settings.Repository)
        Dim p As New PaymentDictionary(DataD, "c:\templates\")
        p.UnReleaseCommFile(FileID)
    End Sub

    Public Function UpdateApplicationStatus(MerchantID As Long, TerminalID As String, NewStatusID As Integer, Optional ApplicationID As Long = 0, Optional ApplicationRef As String = "", Optional MobileNumber As String = "", Optional IDNumber As String = "") As SwitchPayIntegration.Models.Response Implements ISwitchPayAPI.UpdateApplicationStatus
        Dim rxsd As New SwitchPayIntegration.Models.Response
        Dim hasMerchant As Boolean = False
        Dim hasIDNumber As Boolean = False
        Dim hasMobile As Boolean = False
        Dim nofilters As Boolean = False
        Dim results As Boolean = True
        Dim amountvalid As Boolean = False
        Dim bankvalid As Boolean = False
        Dim mobilevalid As Boolean = True
        Dim idvalid As Boolean = True
        Dim alreadycancelled As Boolean = False
        If (Not IsNumeric(IDNumber)) And (IDNumber <> String.Empty) Then
            idvalid = False
        End If
        If (Not IsNumeric(MobileNumber)) And (MobileNumber <> String.Empty) Then
            mobilevalid = False
        End If
        Try
            If (ApplicationID = 0) And (ApplicationRef = String.Empty) And (IDNumber = String.Empty) And (MobileNumber = String.Empty) Then
                nofilters = True
            End If
            Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            Dim rdb As New RulesDBDataContext(My.Settings.RulesDB)
            Merchant = (From q In rdb.Merchants Where q.ID = CLng(MerchantID) Select q).First()
            hasMerchant = True

            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            Try
                If ApplicationID <> 0 Then
                    App = (From q In db.Applications Where q.ID = CLng(ApplicationID) Select q).First()
                ElseIf ApplicationRef <> String.Empty Then
                    App = (From q In db.Applications Where q.Reference = CStr(ApplicationRef) Select q Order By q.DateCreated Descending).First()
                ElseIf IDNumber <> String.Empty Then
                    App = (From q In db.Applications Where q.IDNumber = CStr(IDNumber) Select q Order By q.DateCreated Descending).First()
                Else
                    'Apps = (From q In db.Applications Where q.entityi = CStr(ApplicationRef) Select q).ToArray
                    results = False
                End If
            Catch
                results = False
            End Try
            hdr = hdt.NewHeaderRow()
            hdr("MerchantID") = MerchantID
            hdr("MerchantRef") = Merchant.Reference
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If results Then
                hdr("ApplicationID") = App.ID
                hdr("ApplicationRef") = App.Reference
                hdr("IDNumber") = App.IDNumber
                hdr("MobileNumber") = App.MobileNumber
                hdr("IsError") = False
                App.InternalAuditTypeID = NewStatusID
                db.SubmitChanges()
                Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr2 = mdt.NewMessagesRow()
                mdr2("Code") = "0"
                mdr2("Message") = "Application ID: " & App.ID & " status successfully updated."
                mdr2("IsError") = False
                mdt.AddMessagesRow(mdr2)

            Else

                hdr("IsError") = True
                If Not results Then
                    Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                    mdr2 = mdt.NewMessagesRow()
                    mdr2("Code") = "2"
                    mdr2("Message") = "No Applications Found."
                    mdr2("IsError") = True
                    mdt.AddMessagesRow(mdr2)
                End If
                If nofilters Then
                    Dim mdr2 As SwitchPayIntegration.Models.Response.MessagesRow
                    mdr2 = mdt.NewMessagesRow()
                    mdr2("Code") = "3"
                    mdr2("Message") = "No Filters Provided."
                    mdr2("IsError") = True
                    mdt.AddMessagesRow(mdr2)
                End If
                If Not mobilevalid Then
                    Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                    mdr3 = mdt.NewMessagesRow()
                    mdr3("Code") = "5"
                    mdr3("Message") = "Mobile Number Invalid"
                    mdr3("IsError") = True
                    mdt.AddMessagesRow(mdr3)

                End If

                If Not idvalid Then
                    Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                    mdr3 = mdt.NewMessagesRow()
                    mdr3("Code") = "5"
                    mdr3("Message") = "ID Number Invalid"
                    mdr3("IsError") = True
                    mdt.AddMessagesRow(mdr3)


                End If
            End If

            hdt.AddHeaderRow(hdr)

        Catch ex As Exception
            Dim hdt = New SwitchPayIntegration.Models.Response.HeaderDataTable
            hdt = rxsd.Tables("Header")
            'Dim db As New DBDataContext(My.Settings.SwitchPayDB)
            'Dim apps = (From q In db.Applications Where appl)
            Dim hdr As SwitchPayIntegration.Models.Response.HeaderRow
            hdr = hdt.NewHeaderRow()
            hdr("IsError") = True
            hdt.AddHeaderRow(hdr)
            Dim fdt As New SwitchPayIntegration.Models.Response.FieldsDataTable
            fdt = rxsd.Tables("Fields")
            Dim mdt As New SwitchPayIntegration.Models.Response.MessagesDataTable
            mdt = rxsd.Tables("Messages")
            If hasMerchant = False Then
                Dim mdr3 As SwitchPayIntegration.Models.Response.MessagesRow
                mdr3 = mdt.NewMessagesRow()
                mdr3("Code") = "4"
                mdr3("Message") = "No Valid Merchant Found."
                mdr3("IsError") = True
                mdt.AddMessagesRow(mdr3)

            End If
            Dim mdr As SwitchPayIntegration.Models.Response.MessagesRow
            mdr = mdt.NewMessagesRow()
            mdr("Code") = "1"
            mdr("Message") = "Exception: " & ex.Message
            mdr("IsError") = True
            mdt.AddMessagesRow(mdr)
        End Try
        Return rxsd
    End Function


    Public Sub UpdateDebitOrders() Implements ISwitchPayAPI.UpdateDebitOrders
        Dim DataD As New DataDictionary(My.Settings.Environment, My.Settings.Repository)
        Dim p As New PaymentDictionary(DataD, "c:\templates\")
        p.GenerateDebitOrders()
    End Sub

    Public Function ValidateID(ByVal IDno As String) As Boolean Implements ISwitchPayAPI.ValidateID
        Dim a As Integer = 0
        For i As Integer = 0 To 5
            a += CInt(IDno.Substring(i * 2, 1))
        Next
        Dim b As Integer = 0
        For i As Integer = 0 To 5
            b = (b * 10) + CInt(IDno.Substring((2 * i) + 1, 1))
        Next
        b *= 2
        Dim c As Integer = 0
        Do
            c += b Mod 10
            b = CInt(Int(b / 10))
        Loop Until b <= 0
        c += a
        Dim d As Integer = 0
        d = 10 - (c Mod 10)
        If d = 10 Then d = 0
        If (d = CInt(IDno.Substring(12, 1))) And (IsDate("19" & IDno.Substring(0, 2) & "/" & IDno.Substring(2, 2) & "/" & IDno.Substring(4, 2)) Or IsDate("20" & IDno.Substring(0, 2) & "/" & IDno.Substring(2, 2) & "/" & IDno.Substring(4, 2))) Then
            Return True
        Else
            Return False
        End If

    End Function 'Validate

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#Region "Variables"

    Public App As Application
    Public dd As DataDictionary
    Public md As MerchantDictionary
    Public merchanttype As String
    Dim Merchant As New Merchant

    Public Saved As Boolean = False
#End Region


End Class
