using Microsoft.Exchange.WebServices.Data;
using SyntheticTransactionsForExchange.DataModels;
using SyntheticTransactionsForExchange.Utilities;
using System;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.Management.Automation;
using System.Text.Json;
using System.Threading;

namespace SyntheticTransactionsForExchange.ExchangeMailflowMonitoring
{
    [Cmdlet("Get", "EWSMonitoringMail")]
    [OutputType(typeof(MailflowMonitoringData))]
    [CmdletBinding()]

    public class GetEWSMonitoringMail : Cmdlet
    {
        /// <summary>
        /// <para type="description">The Url of the server to target with the Exchange Web Service connection. If missing autodiscover will be used.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"Exchange Web Services URL")]
        public String Url;

        /// <summary>
        /// <para type="description">The the Exchange Web Service schema version. If missing autodiscover will locate the highest supported version.</para>
        /// <para type="description">If omitted Exchange2013_SP1 will be assumed as default.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"Exchange Web Services Version")]
        public ExchangeVersion Version = ExchangeVersion.Exchange2013_SP1;

        /// <summary>
        /// <para type="description">The mailbox to target with this EWS Service. This parameter is mandatory.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"Mailbox to access via Exchange Web Services")]
        public String Mailbox;

        /// <summary>
        /// <para type="description">The mailbox to target with this EWS Service. This parameter is mandatory.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"Username to access the Mailbox via Exchange Web Services")]
        public String UserName;

        /// <summary>
        /// <para type="description">The mailbox to target with this EWS Service. This parameter is mandatory.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"Password to access the Mailbox via Exchange Web Services")]
        public String Password;

        /// <summary>
        /// <para type="description">The Credentials of the user connecting to the service.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"The Credentials of the user connecting to the service.")]
        public PSCredential Credentials = null;

        /// <summary>
        /// <para type="description">JWT Token to be used for accessing the Mailbox.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"JWT Token to be used for accessing the Mailbox")]
        public String AccessToken;

        /// <summary>
        /// <para type="description">If specified then Application Impersonation would be set; the mailbox impersonated is defined by the attribute "Mailbox".</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"f specified then Application Impersonation would be set; the mailbox impersonated is defined by the attribute Mailbox")]
        public SwitchParameter ApplicationImpersonation;

        /// <summary>
        /// <para type="description">Guid to be searched for in the Subject of the Message.</para>
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = @"Guid to be searched for in the Subject of the Message.")]
        public Guid SubjectGuid;

        /// <summary>
        /// <para type="description">How long to wait, at most, for the message to arrive (in seconds).</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"How long to wait, at most, for the message to arrive (in seconds).")]
        public UInt16 TimeOut = 300;

        /// <summary>
        /// <para type="description">How long to pause between each attempt to fetch data from the mailbox (in seconds).</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"How long to pause between each attempt to fetch data from the mailbox (in seconds).")]
        public UInt16 SleepTimer = 5;

        /// <summary>
        /// <para type="description">If specified EWS/AutoDiscover tracing will be enabled. This is very verbose but offers a great way to see the SOAP for each operation.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"If specified tracing is enabled")]
        public SwitchParameter Trace;

        protected override void BeginProcessing()
        {
            WriteVerbose(String.Format("Entering cmdlet on <{0}>", DateTime.Now.ToString("yyyy-MM-MMTHH:mm:ss")));
        }

        protected override void ProcessRecord()
        {
            if (!String.IsNullOrEmpty(AccessToken) && String.IsNullOrEmpty(Mailbox))
            {
                throw new Exception("Mailbox cannot be null if using OAuth");
            }
            if (String.IsNullOrEmpty(Mailbox) && !String.IsNullOrEmpty(UserName))
            {
                Mailbox = UserName;
                WriteVerbose(String.Format("Setting Mailbox to match Username <{0}>", UserName));
            }
            if (String.IsNullOrEmpty(Mailbox) && Credentials != null)
            {
                Mailbox = Credentials.UserName;
                WriteVerbose(String.Format("Setting Mailbox to match Credentials Username <{0}>", Credentials.UserName));
            }
            if (String.IsNullOrEmpty(Mailbox) && String.IsNullOrEmpty(AccessToken) && Credentials == null && String.IsNullOrEmpty(UserName))
            {
                Mailbox = UserPrincipal.Current.EmailAddress;
                WriteVerbose(String.Format("Setting Mailbox to the SMTP of the logged on user which is <{0}>", Mailbox));
            }

            if ((!String.IsNullOrEmpty(UserName) && String.IsNullOrEmpty(Password)) || (String.IsNullOrEmpty(UserName) && !String.IsNullOrEmpty(Password)))
            {
                throw new Exception("If provided, Username and Password must be both set");
            }

            if (Credentials != null)
            {
                UserName = Credentials.UserName;
                Password = Credentials.GetNetworkCredential().Password;
                WriteVerbose(String.Format("Username and Password set from Credentials"));
            }

            if (!RegexUtilities.IsValidEmail(Mailbox))
            {
                throw new Exception("Mailbox address is invalid");
            }

            if (SubjectGuid == null || SubjectGuid == Guid.Empty)
            {
                throw new Exception("Guid cannot be null");
            }

            MailflowMonitoringData monitoringData = new MailflowMonitoringData();
            ExchangeService EWSService = new ExchangeService(Version);

            try
            {
                if (!String.IsNullOrEmpty(UserName) && !String.IsNullOrEmpty(Password))
                {
                    EWSService.Credentials = new WebCredentials(UserName, Password);
                    WriteVerbose(String.Format("Using Username and Password for <{0}>", UserName));
                }
                else if (!String.IsNullOrEmpty(AccessToken))
                {
                    EWSService.Credentials = new OAuthCredentials(AccessToken);
                    WriteVerbose(String.Format("Using OAuth token <{0}>", AccessToken));
                }
                else
                {
                    EWSService.UseDefaultCredentials = true;
                    WriteVerbose(String.Format("Using Windows Integrated with the currently logged on user which is <{0}>", UserPrincipal.Current.UserPrincipalName));
                }

                if (!String.IsNullOrEmpty(Url))
                {
                    EWSService.Url = new Uri(Url);
                }
                else
                {
                    WriteVerbose(String.Format("Trying to AutoDiscover URL for <{0}>", Mailbox));
                    EWSService.AutodiscoverUrl(Mailbox, delegate { return true; });
                    WriteVerbose(String.Format("The EWS URL has been located and is <{0}>", EWSService.Url.ToString()));
                }

                if (ApplicationImpersonation)
                {
                    EWSService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, Mailbox);
                }
                EWSService.HttpHeaders.Add("X-AnchorMailbox", Mailbox);

                if (Trace)
                {
                    EWSService.TraceEnabled = true;
                    EWSService.TraceFlags = TraceFlags.All;
                }
            }
            catch (Exception ex)
            {
                WriteVerbose(String.Format("Data................: {0}", ex.Data));
                WriteVerbose(String.Format("HelpLink............: {0}", ex.HelpLink));
                WriteVerbose(String.Format("HResult.............: {0}", ex.HResult));
                WriteVerbose(String.Format("InnerException......: {0}", ex.InnerException));
                WriteVerbose(String.Format("Message.............: {0}", ex.Message));
                WriteVerbose(String.Format("Source..............: {0}", ex.Source));
                throw ex;
            }

            DateTime receivedDateTime = new DateTime();
            bool messageReceived = false;
            Stopwatch timer = new Stopwatch();

            try
            {
                Mailbox mbx = new Mailbox(Mailbox);
                FolderId inboxId = new FolderId(WellKnownFolderName.Inbox, mbx);
                Folder inbox = Folder.Bind(EWSService, inboxId);

                ItemView view = new ItemView(500);
                view.PropertySet = PropertySet.FirstClassProperties;
                view.Traversal = ItemTraversal.Shallow;

                PropertySet properySet = new PropertySet(BasePropertySet.FirstClassProperties);
                properySet.Add(EmailMessageSchema.Subject);
                properySet.Add(EmailMessageSchema.Body);
                properySet.Add(EmailMessageSchema.TextBody);
                properySet.Add(EmailMessageSchema.DateTimeSent);
                properySet.Add(EmailMessageSchema.DateTimeReceived);
                properySet.RequestedBodyType = BodyType.Text;

                SearchFilter filter = new SearchFilter.ContainsSubstring(ItemSchema.Subject, SubjectGuid.ToString());

                timer.Start();

                while (messageReceived == false && timer.ElapsedMilliseconds <= TimeOut * 1000)
                {
                    Stopwatch operationTimer = new Stopwatch();
                    operationTimer.Start();
                    FindItemsResults<Item> messages = EWSService.FindItems(inboxId, filter, view);
                    operationTimer.Stop();
                    foreach (Item message in messages)
                    {
                        WriteVerbose(String.Format("Parsing message <{0}> sent from <{1}>", message.Subject, ((EmailMessage)message).Sender));
                        if (message.Subject.Contains(SubjectGuid.ToString()))
                        {
                            messageReceived = true;
                            message.Load(properySet);
                            receivedDateTime = message.DateTimeReceived.ToUniversalTime();
                            monitoringData.SetDuration(Convert.ToUInt32(operationTimer.ElapsedMilliseconds));
                            MailflowMonitoringData temp = JsonSerializer.Deserialize<MailflowMonitoringData>(message.TextBody);
                            monitoringData.SetSendingInformation(temp.SubjectGuid, temp.TimeSent, temp.SendingProtocol);
                            string RAWHeader = String.Join(System.Environment.NewLine, message.InternetMessageHeaders);
                            WriteVerbose(String.Format("The header parsed is <{0}> ", RAWHeader));
                            monitoringData.MailflowHeaderDataTable = MailUtilities.ParseMailHeader(RAWHeader);
                            monitoringData.SetStatus(TransactionStatus.Success);
                            message.Delete(DeleteMode.HardDelete);
                        }
                    }
                    if (!messageReceived)
                    {
                        WriteVerbose(String.Format("Sleeping <{0}> seconds", SleepTimer));
                        Thread.Sleep(SleepTimer * 1000);
                    }
                }
                if (!messageReceived)
                {
                    WriteWarning("The message has not been found!");
                }
            }
            catch (Exception ex)
            {
                WriteVerbose(String.Format("Data................: {0}", ex.Data));
                WriteVerbose(String.Format("HelpLink............: {0}", ex.HelpLink));
                WriteVerbose(String.Format("HResult.............: {0}", ex.HResult));
                WriteVerbose(String.Format("InnerException......: {0}", ex.InnerException));
                WriteVerbose(String.Format("Message.............: {0}", ex.Message));
                WriteVerbose(String.Format("Source..............: {0}", ex.Source));
                throw ex;
            }

            monitoringData.SetReceivingInformation(SubjectGuid, receivedDateTime, Protocol.EWS);

            if (messageReceived)
            {
                monitoringData.ComputeLatency();
            }

            WriteObject(monitoringData);
            WriteVerbose(monitoringData.ToString());
            return;
        }

        protected override void EndProcessing()
        {
            WriteVerbose(String.Format("Exiting cmdlet on <{0}>", DateTime.Now.ToString("yyyy-MM-MMTHH:mm:ss")));
        }
    }
}
