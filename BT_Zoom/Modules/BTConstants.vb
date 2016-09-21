Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls

Namespace Constants

    Public Class BTConstants

        Public Const EmptyInteger As Integer = -999
        Public Shared ReadOnly BTMinDate As DateTime = New DateTime(2000, 1, 1, 0, 0, 0)
        Public Const KeywordSearchParameterLength As Integer = 25
        Public Const KeywordFullTextSearchParameterLength As Integer = 28

#Region "HttpContext Information"


        ' This is additional markup to be assigned to any html <a> tag that uses href="mailto:...". Due to a bug with Chrome,
        ' SSL anchors need the additional target markup to work correctly
        ' Bug Ref: bt-4487

        Public Shared Function IsValidDateString(ByVal dateStr As String) As Boolean
            Dim result As Boolean = False
            Dim tmpDt As DateTime
            If Not String.IsNullOrWhiteSpace(dateStr) AndAlso DateTime.TryParse(dateStr, tmpDt) Then
                result = True
            End If
            Return result
        End Function

        ''' <summary>
        ''' p=1, ret=1
        ''' p=2, ret=2
        ''' p=3, ret=4
        ''' p=4, ret=8
        ''' p=5, ret=16
        ''' </summary>
        ''' <param name="i"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function get2Power(ByVal i As Integer) As Long
            Dim curVal As Int64 = 1
            If i > 1 Then
                curVal = Convert.ToInt64(2 ^ (i - 1))
            End If

            Return (curVal)
        End Function

        Public Shared Function IsHex(ByVal str As String) As Boolean
            If String.IsNullOrWhiteSpace(str) Then _
                Return False

            Dim i As Int32, c As Char

            If str.IndexOf("0x") = 0 Then _
                str = str.Substring(2)

            While (i < str.Length)
                c = str.Chars(i)

                If Not (((c >= "0"c) AndAlso (c <= "9"c)) OrElse
                        ((c >= "a"c) AndAlso (c <= "f"c)) OrElse
                        ((c >= "A"c) AndAlso (c <= "F"c))) _
                    Then
                    Return False
                Else
                    i += 1
                End If
            End While

            Return True
        End Function


        Public Shared Function GetServerIPAddressCollection() As IPAddress()
            Dim ipEntry As IPHostEntry = Dns.GetHostEntry(Dns.GetHostName())
            Dim ipAddr As IPAddress() = ipEntry.AddressList
            Return ipAddr
        End Function

        Public Shared Function GetPostbackControlID(ByVal p As Page) As String
            Dim pcontrol As Control = Nothing
            Dim cname As String
            cname = p.Request.Params("__EVENTTARGET")
            If (cname <> Nothing And cname <> String.Empty) Then
                'pcontrol = p.FindControl(cname)
            Else
                Dim i As String
                For Each i In p.Request.Form
                    If i.EndsWith(".x") Or i.EndsWith(".y") Then
                        'image button
                        cname = i.Substring(0, i.Length - 2)
                        If cname.IndexOf("$") = -1 And cname.IndexOf(":") = -1 Then
                            pcontrol = p.FindControl(cname)
                        End If
                    Else
                        'all other controls & buttons
                        cname = i
                        If cname.IndexOf("$") = -1 And cname.IndexOf(":") = -1 Then
                            pcontrol = p.FindControl(cname)
                        End If
                    End If
                    If (TypeOf pcontrol Is Button) Or (TypeOf pcontrol Is ImageButton) Then Exit For 'we found the button control now quit
                Next
            End If
            Return (cname)
        End Function

#End Region

#Region "Mobile Checks"

#End Region

#Region "Measure String in Pixels"
        Public Shared Function GetMaxStringLengthInPixels(ddl1 As DropDownList, fontName As String, fontSizeInPx As Integer) As Integer
            Using objBitmap As New Bitmap(500, 200)
                Using objGraphics As Graphics = Graphics.FromImage(objBitmap)

                    'http://stackoverflow.com/questions/5553965/how-to-programmatically-measure-string-in-asp-net
                    Dim stringSizeInPixels As SizeF
                    Dim maxWidth As Integer = -1
                    For Each item As ListItem In ddl1.Items
                        ' Note: This is inside of a loop doing the same thing as the string method, due to memory.  we don't want to create an instance of the objBitmap every time inside of a loop calling another method
                        Using f As New Font(fontName, fontSizeInPx, GraphicsUnit.Pixel)
                            stringSizeInPixels = objGraphics.MeasureString(item.Text, f)
                            If maxWidth < stringSizeInPixels.Width Then
                                maxWidth = CInt(stringSizeInPixels.Width)
                            End If
                        End Using
                    Next
                    Return maxWidth

                End Using
            End Using
        End Function

        Public Shared Function GetStringLengthInPixels(str As String, fontName As String, fontSize As Integer) As Integer
            Using objBitmap As New Bitmap(500, 200)
                Using objGraphics As Graphics = Graphics.FromImage(objBitmap)
                    'http://stackoverflow.com/questions/5553965/how-to-programmatically-measure-string-in-asp-net
                    Dim stringSizeInPixels As SizeF
                    Dim MaxWidth As Integer = -1
                    Using f As New Font(fontName, fontSize)
                        stringSizeInPixels = objGraphics.MeasureString(str, f)
                        If MaxWidth < stringSizeInPixels.Width Then
                            MaxWidth = CInt(stringSizeInPixels.Width)
                        End If

                        Return MaxWidth
                    End Using
                End Using
            End Using
        End Function
#End Region

        Public Shared Function GetDateFormatByCulture(culture As String, Optional forMobilePicker As Boolean = False) As String
            Dim format As String
            Select Case culture
                Case "fr-CA", "lv-LV", "lv-EURO"
                    If forMobilePicker Then
                        format = "%Y-%m-%d"
                    Else
                        format = "YYYY-mm-dd"
                    End If

                Case "en-AU", "en-NZ", "en-CA", "en-GB", "ar-EG", "es-ES", "en-ES", "en-CO", "es-CO", "en-NG", "en-NI", "en-ID", "es-PR", "en-NO", "en-IN", "en-DE", "en-TH"
                    If forMobilePicker Then
                        format = "%d-%m-%Y"
                    Else
                        format = "dd-mm-YYYY"
                    End If

                Case "en-US", "en-GH", "en-EG", "en-PG", "en-PH", "en-KE", "en-BS"
                    If forMobilePicker Then
                        format = "%m-%d-%Y"
                    Else
                        format = "mm-dd-YYYY"
                    End If
                Case Else
                    If forMobilePicker Then
                        format = "%m-%d-%Y"
                    Else
                        format = "mm-dd-YYYY"
                    End If

            End Select

            Return format
        End Function

        Public Shared Function IsBTBusinessHours() As Boolean
            If Today.DayOfWeek = DayOfWeek.Saturday OrElse Today.DayOfWeek = DayOfWeek.Sunday Then
                Return False
            End If
            If Now.TimeOfDay.Hours >= 8 AndAlso Now.TimeOfDay.Hours < 17 Then
                Return True
            End If
            Return False
        End Function

        Public Shared Sub IsValidLessEqual(ByVal cuv As CustomValidator, ByVal args As ServerValidateEventArgs, ByVal newValue As Decimal)

            Dim str As String = cuv.Attributes("data-compare")
            If String.IsNullOrWhiteSpace(str) Then
                str = "0"
            End If
            Dim cmp As Decimal
            If Not Decimal.TryParse(str, cmp) Then
                args.IsValid = False
                Exit Sub
            End If

            If newValue > cmp Then
                args.IsValid = False
                Exit Sub
            End If

            args.IsValid = True
        End Sub

        Public Shared Sub IsValidEqual(ByVal cuv As CustomValidator, ByVal args As ServerValidateEventArgs, ByVal newValue As Decimal)

            Dim str As String = cuv.Attributes("data-compare")
            If String.IsNullOrWhiteSpace(str) Then
                str = "0"
            End If
            Dim cmp As Decimal
            If Not Decimal.TryParse(str, cmp) Then
                args.IsValid = False
                Exit Sub
            End If

            args.IsValid = newValue.Equals(cmp)
        End Sub

        Public Shared Sub IsValidGreaterEqual(ByVal cuv As CustomValidator, ByVal args As ServerValidateEventArgs, ByVal newValue As Decimal)

            Dim str As String = cuv.Attributes("data-compare")
            If String.IsNullOrWhiteSpace(str) Then
                str = "0"
            End If
            Dim cmp As Decimal
            If Not Decimal.TryParse(str, cmp) Then
                args.IsValid = False
                Exit Sub
            End If

            If newValue < cmp Then
                args.IsValid = False
                Exit Sub
            End If

            args.IsValid = True
        End Sub

        Public Shared Function GetGoogleMapsImageLink(ByRef linkText As String, ByRef style As String, ByRef className As String) As String
            Return String.Format("<a href='/images/tradeshow/googleMaps.png' target='_blank' class='{0}' style='{1}'>{2}</a>", className, style, linkText)
        End Function

        Public Const GoogleMapsMappedImage As String = "/images/tradeshow/googleMaps2.png"
        Public Const GoogleMapsMappedTabImage As String = "/images/tradeshow/googleMapsTab.png"


        ' Largely based on http://msdn.microsoft.com/en-us/library/ms254503.aspx
        ' NOTE: need the following section in web.config in order to support the last 4 counters
        '<system.diagnostics>
        '  <switches>
        '    <add name="ConnectionPoolPerformanceCounterDetail" value="4" />
        '  </switches>
        '</system.diagnostics>
        Public Class PerformanceCounters
            Private Const NumberOfPerfCounters As Integer = 13
            Private PerfCounters(NumberOfPerfCounters) As Diagnostics.PerformanceCounter

            Public Enum ADO_Net_Performance_Counters
                NumberOfActiveConnectionPools
                NumberOfReclaimedConnections
                HardConnectsPerSecond
                HardDisconnectsPerSecond
                NumberOfActiveConnectionPoolGroups
                NumberOfInactiveConnectionPoolGroups
                NumberOfInactiveConnectionPools
                NumberOfNonPooledConnections
                NumberOfPooledConnections
                NumberOfStasisConnections
                ' The following performance counters are more expensive to track and need to be enabled in the web.config.
                SoftConnectsPerSecond
                SoftDisconnectsPerSecond
                NumberOfActiveConnections
                NumberOfFreeConnections
            End Enum

            Private Sub New()
                SetUpPerformanceCounters()
            End Sub

            Private Sub SetUpPerformanceCounters()
                Me.PerfCounters(NumberOfPerfCounters) = New Diagnostics.PerformanceCounter()

                Dim instanceName As String = GetInstanceName()
                Dim apc As Type = GetType(ADO_Net_Performance_Counters)
                Dim i As Integer = 0
                Dim s As String = ""
                For Each s In [Enum].GetNames(apc)
                    Me.PerfCounters(i) = New Diagnostics.PerformanceCounter()
                    Me.PerfCounters(i).CategoryName = ".NET Data Provider for SqlServer"
                    Me.PerfCounters(i).CounterName = s
                    Me.PerfCounters(i).InstanceName = instanceName
                    i = (i + 1)
                Next
            End Sub

            Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Integer

            Private Function GetInstanceName() As String
                Dim instanceName As String = AppDomain.CurrentDomain.FriendlyName.ToString.Replace("(", "[").Replace(")", "]").Replace("#", "_").Replace("/", "_").Replace("\\", "_")
                Dim pid As String = GetCurrentProcessId.ToString()
                instanceName = (instanceName + ("[" & (pid & "]")))
                'Console.WriteLine("Instance Name: {0}", instanceName)
                'Console.WriteLine("---------------------------")
                Return instanceName
            End Function

            Public Shared Function GetSnapshot(Optional ByVal forHtml As Boolean = False) As String
                Dim pc As New PerformanceCounters()

                Dim sb As New StringBuilder()
                For Each p As Diagnostics.PerformanceCounter In pc.PerfCounters
                    sb.AppendFormat("{0} = {1}{2}", p.CounterName, p.NextValue, If(forHtml, "<br />", ""))
                    sb.AppendLine()
                Next
                Return sb.ToString()
            End Function

            Public Shared Function GetSnapshotList() As List(Of Diagnostics.PerformanceCounter)
                Dim result As New List(Of Diagnostics.PerformanceCounter)

                Dim pc As New PerformanceCounters()
                For Each p As Diagnostics.PerformanceCounter In pc.PerfCounters
                    result.Add(p)
                Next

                Return result
            End Function

            Public Shared Function CompareSnapshots(ByVal snapshot1 As List(Of Diagnostics.PerformanceCounter), ByVal snapshot2 As List(Of Diagnostics.PerformanceCounter)) As List(Of Boolean)
                Dim result As New List(Of Boolean)

                For i As Integer = 0 To snapshot1.Count - 1
                    result.Add(snapshot1(i).NextValue() = snapshot2(i).NextValue())
                Next

                Return result
            End Function
        End Class

        Public Const BTRequestDictionaryFirstRow As String = "firstRow"
        Public Const BTRequestDictionaryLastRow As String = "lastRow"
        Public Const BTRequestDictionaryPreviousID As String = "previousID"

        Public Const TemplateDirectoryPath As String = "btadmin\TemplatedSetup"
        Public Const ConfirmedStatusText As String = "Confirmed"
        Public Const DeclinedStatusText As String = "Declined"

        Public Const PasswordString As String = "PASSWORD IS SET"

        Public Const WebApiPayloadUnavailable As Integer = -1

        Public Const MerchantPaymentCannotDeleteInvoice As String = "There is an online payment associated with this invoice, so it cannot be deleted.  Please contact support to have this removed manually."
        Public Const MerchantPaymentCannotDeleteInvoiceMultiple As String = "There is an online payment associated with one or more invoices, so they cannot be deleted.  Please contact support to have this removed manually."
        Public Const MerchantPaymentPaymentInProgress As String = "This payment is in progress and cannot be deleted at this time. Please try again later."

        Public Const DefaultListItemValue As Integer = -1


        Public Const PortalTypeHeader As String = "PortalType"
        Public Const PortalTypeQs As String = "loginType"
        Public Const IsFirstLoginQs As String = "isFirstLogin"


        Public Const FileUploadHeader As String = "X-BuilderTREND-FileUpload"

        Public Const BTBillingEmail As String = "billing@buildertrend.com"

        Public Const UptimeCheckerQueryString As String = "uptimeChecker"

        Public Const DefaultMaxRowsThreshold As Integer = 5000


        Public Const MaxAllowedFileSize As Integer = 31457280 ' 30 MB

        Public Const ConstBTAdminDBName As String = "BTAdminDBName"

        Public Class RegexStrings
            'Public Const EmailRegex As String = "(?i:((^((([a-z]|[0-9]|!|#|$|%|&|'|\*|\+|\-|/|=|\?|\^|_|`|\{|\||\}|~)+(\.([a-z]|[0-9]|!|#|$|%|&|'|\*|\+|\-|/|=|\?|\^|_|`|\{|\||\}|~)+)*)@((((([a-z]|[0-9])?([a-z]|[0-9]|\-){0,61}([a-z]|[0-9])\.))*([a-z]|[0-9])([a-z]|[0-9]|\-){0,61}([a-z]|[0-9])?(\.(af|ax|al|dz|as|ad|ao|ai|aq|ag|ar|am|aw|au|at|az|bs|bh|bd|bb|by|be|bz|bj|bm|bt|bo|ba|bw|bv|br|io|bn|bg|bf|bi|kh|cm|ca|cv|ky|cf|td|cl|cn|cx|cc|co|km|cg|cd|ck|cr|ci|hr|cu|cy|cz|dk|dj|dm|do|ec|eg|sv|gq|er|ee|et|fk|fo|fj|fi|fr|gf|pf|tf|ga|gm|ge|de|gh|gi|gr|gl|gd|gp|gu|gt| gg|gn|gw|gy|ht|hm|va|hn|hk|hu|is|in|id|ir|iq|ie|im|il|it|jm|jp|je|jo|kz|ke|ki|kp|kr|kw|kg|la|lv|lb|ls|lr|ly|li|lt|lu|mo|mk|mg|mw|my|mv|ml|mt|mh|mq|mr|mu|yt|mx|fm|md|mc|mn|ms|ma|mz|mm|na|nr|np|nl|an|nc|nz|ni|ne|ng|nu|nf|mp|no|om|pk|pw|ps|pa|pg|py|pe|ph|pn|pl|pt|pr|qa|re|ro|ru|rw|sh|kn|lc|pm|vc|ws|sm|st|sa|sn|cs|sc|sl|sg|sk|si|sb|so|za|gs|es|lk|sd|sr|sj|sz|se|ch|sy|tw|tj|tz|th|tl|tg|tk|to|tt|tn|tr|tm|tc|tv|ug|ua|ae|gb|uk|us|um|uy|uz|vu|ve|vn|vg|vi|wf|eh|ye|zm|zw|com|edu|gov|int|mil|net|org|biz|info|name|pro|aero|coop|museum|arpa|asia|eu))+)|(((([0-9]){1,3}\.){3}([0-9]){1,3}))|(\[((([0-9]){1,3}\.){3}([0-9]){1,3})\])))\b)(;\s*(((([a-z]|[0-9]|!|#|$|%|&|'|\*|\+|\-|/|=|\?|\^|_|`|\{|\||\}|~)+(\.([a-z]|[0-9]|!|#|$|%|&|'|\*|\+|\-|/|=|\?|\^|_|`|\{|\||\}|~)+)*)@((((([a-z]|[0-9])?([a-z]|[0-9]|\-){0,61}([a-z]|[0-9])\.))*([a-z]|[0-9])([a-z]|[0-9]|\-){0,61}([a-z]|[0-9])?(\.(af|ax|al|dz|as|ad|ao|ai|aq|ag|ar|am|aw|au|at|az|bs|bh|bd|bb|by|be|bz|bj|bm|bt|bo|ba|bw|bv|br|io|bn|bg|bf|bi|kh|cm|ca|cv|ky|cf|td|cl|cn|cx|cc|co|km|cg|cd|ck|cr|ci|hr|cu|cy|cz|dk|dj|dm|do|ec|eg|sv|gq|er|ee|et|fk|fo|fj|fi|fr|gf|pf|tf|ga|gm|ge|de|gh|gi|gr|gl|gd|gp|gu|gt| gg|gn|gw|gy|ht|hm|va|hn|hk|hu|is|in|id|ir|iq|ie|im|il|it|jm|jp|je|jo|kz|ke|ki|kp|kr|kw|kg|la|lv|lb|ls|lr|ly|li|lt|lu|mo|mk|mg|mw|my|mv|ml|mt|mh|mq|mr|mu|yt|mx|fm|md|mc|mn|ms|ma|mz|mm|na|nr|np|nl|an|nc|nz|ni|ne|ng|nu|nf|mp|no|om|pk|pw|ps|pa|pg|py|pe|ph|pn|pl|pt|pr|qa|re|ro|ru|rw|sh|kn|lc|pm|vc|ws|sm|st|sa|sn|cs|sc|sl|sg|sk|si|sb|so|za|gs|es|lk|sd|sr|sj|sz|se|ch|sy|tw|tj|tz|th|tl|tg|tk|to|tt|tn|tr|tm|tc|tv|ug|ua|ae|gb|uk|us|um|uy|uz|vu|ve|vn|vg|vi|wf|eh|ye|zm|zw|com|edu|gov|int|mil|net|org|biz|info|name|pro|aero|coop|museum|arpa|asia|eu))+)|(((([0-9]){1,3}\.){3}([0-9]){1,3}))|(\[((([0-9]){1,3}\.){3}([0-9]){1,3})\])))\b))*);?\s*$)"
            Public Const CellEmailRegex As String = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
            Public Const FriendlyEmailName As String = "(?=^[A-Za-z])(^[A-Za-z 0-9,.\-&']+$)"

            Public Shared Function FriendlyEmailNameWithLength(ByVal minval As Integer, ByVal maxval As Integer) As String
                Return String.Format("(?=^[A-Za-z])(^[A-Za-z 0-9,.\-&']{{{0},{1}}}$)", minval, maxval)
            End Function

            Private Const EmailBase As String = "\s*([a-zA-Z0-9]([\.]?(\w+|[!#$%&'\*\+\-/=\?\^`\{\|\}~]+))*)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}(\]))|(((\w+([!#$%&'\*\+\-/=\?\^`\{\|\}~])*)*\w+(\.))+)([a-zA-Z0-9_]+[a-zA-Z]))\s*"

            'not technically constants but are regex strings so leaving them here
            'I did it this way so that if we need to edit the way we regex email, we don't have to change it in several places, just the EmailBase
            ''' <summary>
            ''' Allows for one email address (with no delimiter)
            ''' </summary>
            Public Shared ReadOnly Property EmailSingle As String
                Get
                    Return String.Format("^{0}$", EmailBase)
                End Get
            End Property

            ''' <summary>
            ''' Allows for one or multiple email addresses delimited by ';' with any number of trailing/leading spaces between addresses, if only one - delimiter is not required
            ''' </summary>
            Public Shared ReadOnly Property EmailRegex As String
                Get
                    Return String.Format("^(({0})({1}{0})*{1}?)$", EmailBase, ";")
                End Get
            End Property
        End Class


        Enum InlineControlType
            DropDown = 1
            TextBox = 2
            DatePicker = 3
            MultiLineTextBox = 4
            DropDownChecklist = 5
        End Enum

        Public Class Bidding
            Public Const BiddingTitle As String = "Bids"

            Public Const BidSingular As String = "Bid Request"
            Public Const BidPlural As String = "Bid Requests"
            Public Const BidSingularOrPlural As String = "Bid Request(s)"

            Public Const BidPackageSingular As String = "Bid Package"
            Public Const BidPackagePlural As String = "Bid Packages"
            Public Const BidPackageSingularOrPlural As String = "Bid Package(s)"

            Public Const BidNotSubmittedText As String = "Did Not Submit"

            Public Class RFIs
                Public Const AssignedToAllUsers As Integer = -10
            End Class

            Public Class QueryStringParamNames
                Public Const BidPackageId As String = "bidPackageId"
                Public Const BidId As String = "bidId"
            End Class

            Public Class MenuLabels
                Public Const ImportBids As String = "Import Bid Packages"
            End Class
        End Class

        Public Class ReloadPayPalPacket

            Public Class QueryStringParamNames
                Public Const FromReloadPayPalPacket As String = "FromReloadPayPalPacket"
            End Class

        End Class

        Public Class Leads
            Public Const LeadActivitySent As String = "Lead Activity Sent"

            Public Class QueryStringParameters
                Public Const LeadActivityID As String = "activityID"
                Public Const LeadID As String = "leadID"
                Public Const ParentDialogID As String = "parentDialog"
            End Class
        End Class

        Public Class Surveys
            Public Const Survey As String = "Survey"
            Public Const Surveys As String = "Surveys"
            Public Const SurveyQuestions As String = "Survey Questions"
            Public Const SurveyDefinitions As String = "Survey Definitions"
            Public Const SurveyQuestionAnswers As String = "Survey Question Answers"
            Public Const SurveyList As String = "Survey List"

            Public Const SurveySent As String = "Survey Sent"
            Public Const SurveyScheduled As String = "Survey Scheduled"
            Public Const SurveyCompleted As String = "Survey Completed"

            Public Class QueryStringParameters
                Public Const SurveyQuestionID As String = "QuestionID"
                Public Const SurveyDefinitionID As String = "surveyDefinitionID"
                Public Const SurveyID As String = "surveyID"
            End Class
        End Class

        Public Class PurchaseOrders
            ' Used for mobile for now since the value in the dropdown for web is "ALL"
            Public Const SubApprovalAllStatuses As Integer = -1

            Public Const ScheduleTitleAlias As String = "SchedTitle"

            Public Const PurchaseOrderSingular As String = "Purchase Order"

            Public Const UnassignedPOTrackPaymentsWarning As String = "In order to track payments on this PO, it must first be assigned."
            Public Const UnassignedPOMarkCompleteWarning As String = "In order to mark a purchase order work complete, it must first be assigned."

            Public Const PaymentStatus_InProgressText As String = "Not Complete"
            Public Const PaymentStatus_WorkCompletedText As String = "Work Complete"
            Public Const PaymentStatus_PaidText As String = "Paid"
            Public Const PaymentStatus_RequestPayment As String = "Payment Requested"
            Public Const PO_Fully_CompleteGreaterThanThis As Integer = 40 'if the poStatus is greater than this, it is completed
            Public Const PO_ApprovalStatus_All As String = "ALL"

            ''' <summary>
            ''' if the poStatus is greater than this, the work has passed inspection
            ''' </summary>
            ''' <remarks></remarks>
            Public Const PO_Work_CompleteGreaterThanThis As Integer = 30


            Public Class MenuLabels
                Public Const ImportPOs As String = "POs From Template"
            End Class

            Public Class QueryStringParameters
                Public Const POID As String = "poid"
                Public Const ShowReviewPayments As String = "review"
            End Class

            Public Class Payments
                Public Const POP_AccountingIDs_TableAlias As String = "acc"
                Public Const POP_AccountingBillID_ColumnPrefix As String = "acc_"
                Public Const PurchaseOrderPaymentTitle As String = "Purchase Order Payment"

                Public Shared Function GetLineAndPaymentsQry(ByVal whereClause As String) As String

                    Dim result As String = String.Format(
                        <string>
                        With poli as (
                        Select pi2.PurchaseOrderId, 
                                ISNULL(SUM(pi2.CalculatedAmount), 0) as CurrentPrice
                          from dimension.PurchaseOrderLineItems pi2
                               INNER JOIN dimension.PurchaseOrders po  on po.purchaseOrderId = pi2.PurchaseOrderId 
                        {0} 
                         group by pi2.PurchaseOrderId 
                        ),
                        LineCounts as (
                        SELECT p3.purchaseOrderId, COUNT(1) as noneZeroAmountLines
                          FROM dimension.PurchaseOrderLineItems p3 
                               INNER JOIN poli on poli.PurchaseOrderId = p3.PurchaseOrderId
                         WHERE p3.CalculatedAmount != 0 
                         group by p3.PurchaseOrderId 
                        ),
                        Payments as (
                        SELECT pmnts.PurchaseOrderId, COUNT(1) as countPayments,
                                ISNULL(SUM(pmnts.Amount),0) as Amount
                          FROM dimension.PurchaseOrderPaymentLineItems pmnts 
                               INNER JOIN poli on poli.PurchaseOrderId = pmnts.PurchaseOrderId 
                               INNER JOIN dimension.PurchaseOrderPayments pop ON pop.PurchaseOrderPaymentId = pmnts.PurchaseOrderPaymentId AND pop.paymentStatus != 3
                         group by pmnts.PurchaseOrderId 
                        ),
                        IndividualLines as (
                        SELECT pli.PurchaseOrderId, pli.PurchaseOrderLineItemId, ISNULL(SUM(pli.Amount),0) as LinePaymentAmount
                          FROM dimension.PurchaseOrderPaymentLineItems pli 
                               INNER JOIN poli on poli.PurchaseOrderId = pli.PurchaseOrderId
                               INNER JOIN dimension.PurchaseOrderPayments pop ON pop.PurchaseOrderPaymentId = pli.PurchaseOrderPaymentId AND pop.paymentStatus != 3
                         group by pli.PurchaseOrderId, pli.PurchaseOrderLineItemId
                        ),
                        LineAndPayment as (
                        select poli.PurchaseOrderId, poli.CurrentPrice, ISNULL(LineCounts.noneZeroAmountLines,0) as countLines, 
                                ISNULL(Payments.amount,0) as Amount, ISNULL(Payments.countPayments,0) as paymentsCount,
                               (select count(1) 
                                  from IndividualLines il 
                                       inner join dimension.PurchaseOrderLineItems pi4 on pi4.PurchaseOrderLineItemId = il.PurchaseOrderLineItemId
                                 where il.PurchaseOrderId = poli.PurchaseOrderId and il.LinePaymentAmount = pi4.CalculatedAmount AND pi4.CalculatedAmount != 0 ) as countFullyPaid
                          from poli
                               left join Payments on poli.PurchaseOrderId = Payments.PurchaseOrderId
                               left join LineCounts on poli.PurchaseOrderId = LineCounts.PurchaseOrderId
                        )    
                    </string>.Value, whereClause
                    )
                    Return result
                End Function

                'NOTE: When changing this, be sure to update the pr_PurcahseOrders_SelectOne and pr_Logins_ValidateDelete stored procs as well
                Public Shared JoinForPaymentAmount As String = "INNER JOIN LineAndPayment on LineAndPayment.PurchaseOrderId = po.purchaseOrderId"

                'NOTE: When changing this, be sure to update the pr_PurcahseOrders_SelectOne and pr_Logins_ValidateDelete stored procs as well
                Public Shared SelectForPaymentAmount As String =
                    <string>
                    LineAndPayment.CurrentPrice as CurrentPrice, 
                    LineAndPayment.Amount as TotalAmountPaid,
                    CASE WHEN LineAndPayment.paymentsCount > 0 THEN 1 ELSE 0 END as HasPayments,
                    CASE WHEN LineAndPayment.countLines = LineAndPayment.countFullyPaid THEN 1 ELSE 0 END AS IsFullyPaid
                </string>.Value

                Public Shared POColumnsList As String =
                    <string>
                    po.purchaseOrderId, po.builderID, po.jobsiteID, po.poName, po.dateAdded, po.addedById, po.dateLastUpdated, po.Title, po.scopeOfWork, 
                    po.internalNotes, po.currentPOStatus, po.performingUserID, po.performingUserType, po.isMaterials, po.estimatedCompletionDate, po.linkedScheduleID, 
                    po.paymentAccountingSystemID, po.lastApprovalRemindedDate, po.lastCompletedRemindedDate, po.lastPaymentReleaseRemindedDate, po.performingUserName, 
                    po.currentPOStatusLastUpdated, po.currentPOStatusLastUpdatedBy, po.isSubApproved_ReadOnly, po.subPartiallyPaid_LockedIn, po.subApprovalStatus, 
                    po.SubApprovalStatusLastUpdated, po.SubApprovalStatusLastUpdatedBy, po.subApprovalComments, po.importedFromPurchaseOrderID, po.tradeAgreementFilePath, 
                    po.paymentCheckNumber, po.SubApprovedPrice, po.isVariance, po.varianceCode, po.relatedPurchaseOrderID, po.relatedChangeOrderID, 
                    po.qb_WasPartiallyPaidAtConversion, po.accountingPartialPaymentIndex, po.signaturePath, po.disclaimerWhenClosed
                </string>.Value


                'NOTE: When changing this, be sure to update the pr_PurcahseOrders_SelectOne and pr_Logins_ValidateDelete stored procs as well
                Public Shared CostCodeSelectForPaymentAmount As String =
                    <string>
                      ISNULL(SUM(poli.CalculatedAmount), 0) as CurrentPrice
                    , ISNULL(SUM(popli.amount), 0) as TotalAmountPaid
                    , CASE WHEN SUM(ISNULL(popli.PurchaseOrderPaymentLineItemId, 0)) > 0 THEN 1 ELSE 0 END as HasPayments
                    , CASE WHEN ISNULL(COUNT(1), 0) = ISNULL(SUM(ISNULL(popli.itemFullyPaid, 0)), 0) THEN 1 ELSE 0 END AS IsFullyPaid
                </string>.Value

                Public Const PaymentDeclinedBySubWarningText As String = "Online payment transaction has been declined by the subcontractor.  Your accounting system will need to be manually updated to reflect this."
                Public Const UnsuccessfulTransactionWarningText As String = "Online payment transaction was unsuccessful.  Your accounting system will need to be manually updated to reflect this."

                Public Class QueryStringParameters
                    Public Const PoPaymentId As String = "POPaymentId"
                    Public Const Check As String = "check"
                    Public Const CreditCard As String = "cc"
                    Public Const redirectUrl As String = "redirectUrl"
                    Public Const MerchantPaymentMethodId As String = "mpmid"
                    Public Const TempFileId As String = "tfid"
                    Public Const PaymentIds As String = "popids"
                    Public Const LienWaiverPaymentId As String = "lwpid"
                    Public Const AlwaysAcceptType As String = "alwaysAccept"
                    Public Const ShowLienWaiver As String = "showLw"
                    Public Const AcceptOrDeclineOnClose As String = "accdecclose"
                End Class

            End Class
        End Class

        Public Class Merchants
            Public Class PaymentMethods
                Public Const StoredPaymentMethodUsageCount As String = "useCount"
            End Class
            Public Const UniqueCheckoutId As String = "uid"
            Public Const OnlinePaymentsNotificationAddresses As String = "steveadmin@buildertrend.com;kkennedy@buildertrend.com;ssiegert@buildertrend.com;jduggeradmin@buildertrend.com;ccaldwell@buildertrend.com;jmarion@buildertrend.com"

            ' Here is a constant we can use for items that are only being sent to colin for now. Then, as more people are added, add their email after the semicolon
            Public Const PaymentConsultantEmailAddresses As String = "ccaldwell@buildertrend.com;"
        End Class

        Public Class ToDos
            Public Const ToDoTitle As String = "To-Do"
            Public Const ToDoSingular As String = "To-Do" 'incase the Menu title should be different
            Public Const ToDoPlural As String = "To-Do's"
            Public Const AllToDos As String = "All To-Do's"
            Public Const AllToDosForBuilder As String = "All Builder Created To-Do's"

            Public Class QueryStringParamNames
                Public Const ToDoId As String = "toDoId"
                Public Const ShowCompletedInfo As String = "showCompleted"
                Public Const HideCompleted As String = "hideCompleted"
                Public Const HideOlderThan30Days As String = "hideOlderThan30Days"
                Public Const QsToDoId As String = "ToDoId"
                Public Const QsFromSummary As String = "s"
                Public Const QsHasBuilder As String = "b"
            End Class

            Public Const AllPriorities As String = "-1"
            Public Const MaxNotesChars As Integer = 100
        End Class

        Public Class TimeClock
            Public Const TimeClockTitle As String = "Time Clock"
            Public Const TimeSheet As String = "Time sheet"

            Public Const PunchIdLookup As String = "{TIMECLOCK_ID}"
            Public Const TabIdLookup As String = "{TAB_ID}"
            Public Const PunchIdAttr As String = "tcId"
            Public Const CurrentTimeUtcAttr As String = "startTime"
            Public Const TimeFormatAttr As String = "timeFormat"
            Public Const OriginalTimeAttr As String = "ot"

            Public Const UnassignedCostCodeId As Integer = -1

            Public Class QueryStringParamNames
                Public Const TimeClockId As String = "TimeClockID"
                Public Const ShiftStatusType As String = "Status"
                Public Const DialogType As String = "EditType"
                Public Const TabToSelect As String = "SubTab"
            End Class

        End Class

        Public Class Documents
            ''' <summary>
            ''' Document instances are grouped into folders that equal their DocumentInstanceType (AssociatedType) * -1000.
            ''' For example, Purchase Order docs will be grouped into a folderId of -6000 (DocumentInstanceType.PurchaseOrder = 6 * -1000)
            ''' </summary>
            ''' <remarks></remarks>
            Public Const DocumentInstanceTypeMultiplier As Integer = -1000
            Public Const AttachedDocumentsFolderId As Integer = -1
            Public Const GlobalDocsFolderId As Integer = -99
            Public Const SubDocumentFolderId As Integer = -2
            Public Const SubDocumentFolderName As String = "** Subcontractor Uploaded Files **"
            Public Const OwnerDocumentFolderId As Integer = -3
            Public Const OwnerDocumentFolderName As String = "** Owner Uploaded Files **"

            Public Class QueryStringParamNames
                Public Const AnnotateDocInstanceID As String = "id"
                Public Const AnnotationID As String = "annotationid"
                Public Const UsesHtml5 As String = "UsesHTML5"
                Public Const DocumentInstanceId As String = "documentInstanceID"
                Public Const AnnotationEncryptedVal As String = "f"
                Public Const UsesAtalasoft As String = "usesAtalasoft"
                Public Const FolderId As String = "folderID"
                Public Const JobsiteId As String = "jobsiteID"
                Public Const DocInstanceIds As String = "docInstances"
                Public Const TempFileId As String = "tfid"
                Public Const Notify As String = "ntfy"
            End Class

        End Class

        Public Class DailyLogs
            Public Const DailyLogSingular As String = "Daily Log"

            Public Const DailyLogTitle As String = "Logs"

            Public Const MapViewLookup As String = "{MAP_VIEW}"

            Public Class QueryStringParamNames
                Public Const DailyLogId As String = "dailyLogId"
                Public Const QsHasBuilder As String = "b"
                Public Const MapView As String = "mapView"
            End Class
        End Class

        Public Class Messages
            ' Global folder does not exist in DB, leave out of Core.Enums.Messages.ReservedEmailFolder enum

            Public Const GlobalEmailFolder As Integer = -9

            ' Search all folder does not exist in DB, leave out of Core.Enums.Messages.ReservedEmailFolder enum

            Public Const SearchAllFolder As Integer = -8

            Public Const MessagesTitle As String = "Messages"

            Public Const ReplyPrefix As String = "RE: "
            Public Const ForwardPrefix As String = "FWD: "

            Public Class UI
                Public Const RecipientStringMaxLength As Integer = 15
            End Class

            Public Class QueryStringParamNames
                Public Const RecipientId As String = "recipientid"
                Public Const MessageId As String = "messageId"
                Public Const DraftId As String = "draftId"
                Public Const Subject As String = "subject"
                Public Const DocsTab As String = "docsTab"
            End Class
        End Class

        Public Class Photos
            Public Const ThumbnailMaxWidth As Integer = 300
            Public Const ThumbnailMaxHeight As Integer = 300
            Public Const PhotoColumn As String = "Photo"
            Public Const RelatedItemColumn As String = "Related Item"
            Public Const AttachedPhotosAlbumId As Integer = -1
            Public Const AttachedPhotoAlbumMultiplier As Integer = -10

            Public Class AttachedPhotosViewModel
                Public Const RelatedItemID As String = "relatedItemID"
                Public Const RelatedItemTitle As String = "relatedItemTitle"
                Public Const RelatedInfo As String = "relatedInfo"
                Public Const Photo As String = "photo"
            End Class

            Public Class QueryStringParamNames
                Public Const AlbumId As String = "albumID"
            End Class
        End Class

        Public Class ScheduleItems
            Public Const IsChangeCol As String = "isChanged"
            Public Const Plural As String = "Schedule Items"
            Public Const Singular As String = "Schedule Item"

            Public Const CheckIdAttribute As String = "id"
            Public Const ConfirmedDivClass As String = "sConfirmed"
            Public Const DeclinedDivClass As String = "sDeclined"
            Public Const PendingDivClass As String = "sPending"

            Public Const PerformedByNameAlias As String = "performedByName"
        End Class

        Public Class ScheduleShifts
            Public Const titleColumnNameFromShiftReasons = "title"
        End Class

        Public Class SelectionItems
            Public Const SelectionSingular As String = "Selection"

            Public Const LocationTitle As String = "locationName"
            Public Const CategoryTitle As String = "categoryName"
            Public Const SchedulesTitle As String = "scheduleTitle"

            Public Class QueryStringParamNames
                Public Const SelectionId As String = "selectionID"
            End Class

            Public Class Allowances
                Public Class QueryStringParamNames
                    Public Const AllowanceId As String = "allowanceID"
                End Class
            End Class

            Public Class Choices
                Public Const ChoiceSingular As String = "Choice"

                Public Class QueryStringParamNames
                    Public Const ChoiceId As String = "choiceID"
                End Class
            End Class
        End Class

        Public Class Builders
            Public Const GettingStartedQuestionCount As Integer = 8 ' Adapted from the readonly property from the old BuilderGettingStarted entity.
            Public Const UNLIMITED_JOBSITES As Integer = 9999

            Public Const BuilderPriceAlias As String = "builderPrice"
            Public Const TotalFeeAlias As String = "totalFee"
            Public Const DaysSinceLastLoginAlias As String = "DaysSinceLastLogin"

            Public Const BuildersAccounting_AccountingIDs_BillAP_TableAlias As String = "abaBillAP_"
            Public Const BuildersAccounting_AccountingIDs_BillAP_ColumnPrefix As String = "abaBillAP_"
            Public Const BuildersAccounting_AccountingIDs_CostCodeExpense_TableAlias As String = "abaCostCodeExpense_"
            Public Const BuildersAccounting_AccountingIDs_CostCodeExpense_ColumnPrefix As String = "abaCostCodeExpense_"
            Public Const BuildersAccounting_AccountingIDs_InvoiceAR_TableAlias As String = "abaInvoiceAR_"
            Public Const BuildersAccounting_AccountingIDs_InvoiceAR_ColumnPrefix As String = "abaInvoiceAR_"
            Public Const BuildersAccounting_AccountingIDs_InvoiceItemID_TableAlias As String = "abaInvoiceItemID_"
            Public Const BuildersAccounting_AccountingIDs_InvoiceItemID_ColumnPrefix As String = "abaInvoiceItemID_"
            Public Const BuildersAccounting_AccountingIDs_CostCodeIncome_TableAlias As String = "abaCostCodeIncome_"
            Public Const BuildersAccounting_AccountingIDs_CostCodeIncome_ColumnPrefix As String = "abaCostCodeIncome_"
            Public Const BuildersAccounting_AccountingIDs_MerchantFee_TableAlias As String = "abaMerchantFee_"
            Public Const BuildersAccounting_AccountingIDs_MerchantFee_ColumnPrefix As String = "abaMerchantFee_"
            Public Const BuildersAccounting_AccountingIDs_MerchantBank_TableAlias As String = "abaMerchantBank_"
            Public Const BuildersAccounting_AccountingIDs_MerchantBank_ColumnPrefix As String = "abaMerchantBank_"

            Public Class QueryStringParamNames
                Public Const BuilderId As String = "builderId"
                Public Const OpenToPaymentTypes As String = "openToPaymentTypes"
            End Class

        End Class

        Public Class Jobs

            Public Const DetailsOpener As String = "jobDetailsOpener"

            Public Const ListOpener As String = "jobsListOpener"

            Public Const InfoTabName As String = "JobInfo"
            Public Const ListTabName As String = "JobsList"
            Public Const NewJobTabName As String = "NewJob"
            Public Const FromScratchTabName As String = "FromScratch"
            Public Const FromTemplateTabName As String = "FromTemplate"
            Public Const TemplateListTabName As String = "TemplatesList"
            Public Const CustomerInvoicesTabName As String = "OwnerPayments"
            Public Const CustomerInvoicesListOpener As String = "customerInvoicesListOpener"

            Public Class QueryStringParamNames
                Public Const JobId As String = "jobId"
                Public Const JobsiteId As String = "jobsiteID"
                Public Const EncryptedJobFilters As String = "jobIDs"
                Public Const JobIds As String = "jobIDs"
                Public Const LeadId As String = "leadID"
                Public Const RefreshOnClose As String = "refreshOnClose"
                Public Const FromGrid As String = "fromGrid"
                Public Const SelectJob As String = "selectJob"
                Public Const FromTemplate As String = "fromTemplate"
                Public Const NewJob As String = "newJob"
                Public Const NewCopyFromTemplate As String = "newTemplCopy"
                Public Const LaunchAccountingLinkCustomer As String = "launchAccountingLinkCustomer"
                Public Const TabId As String = "tbId"
            End Class


            Public Const NotOnJobCssClass As String = " ui-chosen-highlight" 'TODO: create this style and rename if desired

            Public Const PleaseSelectAJob As String = "-- Please Select a Jobsite --"
            Public Const PleaseSelectAJobMobile As String = "Please Select a Jobsite"
            Public Const ChooseATemplate As String = "-- Choose a Template --"
            Public Const ChooseAJob As String = "-- Choose a Jobsite --"
            Public Const ADMIN_PREFIX As String = "**ADMIN** - "
            Public Const JS_Customer_AccountingIDs_TableAlias As String = "custID"
            Public Const JS_Customer_AccountingIDs_ColumnPrefix As String = "customer_"
            Public Const JS_Job_AccountingIDs_TableAlias As String = "jobID"
            Public Const JS_Job_AccountingIDs_ColumnPrefix As String = "job_"
            Public Const Jobsites_AccountingIDs_CustomerID_TableAlias As String = "ajobcust"
            Public Const Jobsites_AccountingIDs_CustomerID_ColumnPrefix As String = "ajobcust_"
            Public Const Jobsites_AccountingIDs_JobID_TableAlias As String = "ajobid"
            Public Const Jobsites_AccountingIDs_JobID_ColumnPrefix As String = "ajobid_"
            Public Const RelScheduleAssignedUserLinks As String = "SchedulesToAssignees"
            Public Const TblLinkedToDos As String = "LinkedToDos"
            Public Const RelToDoScheduleLinks As String = "ToDoLinkedSchedule"
            Public Const TblLinkedCustomerInvoices As String = "LinkedCustomerInvoices"
            Public Const RelCustomerInvoicesScheduleLinks As String = "CustomerInvoicesLinkedSchedule"
            Public Const POTable As String = "PurchaseOrders"
            Public Const PaymentsTable As String = "Payments"
            Public Const POPaymentsRelationship As String = "PORelatedPayments"
            Public Const ScheduleLinkedPaymentsRelationship As String = "ScheduleLinkedPayments"
        End Class

        Public Class Subs
            Public Const Sub_AccountingIDs_TableAlias As String = "asub"
            Public Const Sub_AccountingIDs_ColumnPrefix As String = "asub_"

            Public Class QueryStringParamNames
                Public Const JobsiteId As String = "jobsiteID"
                Public Const SubId As String = "subId"
                Public Const SubIds As String = "subIDs"
                Public Const JobPickerBuilderId As String = "jpbid"
            End Class
        End Class

        Public Class PaymentInfo
            Public Const Name As String = "Name"
            Public Const CompanyName As String = "CompanyName"
            Public Const Email As String = "Email"
        End Class

        Public Class DynamicGrid

            Public Const SortColumn As String = "sortCol"

            Public Const SortDirection As String = "sortDir"

            Public Const IsSystemView As String = "SystemView"

            Public Const NonSortableColumn As String = "NonSortableColumn"
        End Class

        Public Class WebAPIAliases
            Public Const IDColumn As String = "id"
            Public Const JobID As String = "jobId"
        End Class

        Public Class Calendar
            Public Const TooManyScheduleItemsThreshold As Integer = 4000
            Public Const ItemSourceId As String = "{BT_SOURCE_ID}"
            Public Const DateParam As String = "{BT_DATE}"
            Public Const HidFilterPostVal As String = "encFilter"
            Public Const HidDatePostVal As String = "calDate"

            Public Class QueryStringParamNames
                Public Const TaskId As String = "taskID"
                Public Const WorkdayExceptionId As String = "WorkdayExceptionID"
                Public Const SelectedTab As String = "selectedTab"
                Public Const ScheduleIds As String = "scheduleIds"
                Public Const ScheduleId As String = "scheduleId"
                Public Const CalendarFeedId As String = "f"
                Public Const TurnCalendarOnline As String = "TurnCalendarOnline"
                Public Const VisitedFrom As String = "VisitedFrom"
            End Class

            Public Class OwnerQueryStringParamNames
                Public Const InitialView As String = "initialView"
                Public Const InitialFilter As String = "initialFilter"
            End Class

            Public Class CalendarMenuLabels
                Public Const ScheduleItem As String = "Schedule Item"
                Public Const Task As String = "Task"
                Public Const WorkDayException As String = "Workday Exception"
                Public Const SetBaseLine As String = "Set Baseline"
                Public Const DeleteAllItems As String = "Delete All Items"
                Public Const ImportCalendar As String = "Import Calendar"
                Public Const NotifyUsers As String = "Notify Assigned Users"
                Public Const TrackCollision As String = "Track Collisions"
                Public Const WorkOnline As String = "Work Online"
                Public Const WorkOffline As String = "Work Offline"
                Public Const PrintWithBrowser As String = "Print with Browser"
                Public Const PrintToPDF As String = "Print to PDF<span style='font-size:10px; margin-left:5px;'>(recommended)</span>"
                Public Const PrintLandScape As String = "Landscape"
                Public Const PrintPortrait As String = "Portrait"
                Public Const DeleteCheckedItems As String = "Delete Checked Items"
                Public Const OwnerViewingPermission As String = "Owner Viewing Permission"
                Public Const DailyLog As String = "Daily Log"
            End Class

        End Class

        Public Class DAL
            Public Const PagingRowIdColumnName As String = "rowId"
            Public Const PagingTotalRowsColumnName As String = "totalRows"
            Public Const RowNumber As String = "rowNumber"
        End Class

    End Class

End Namespace