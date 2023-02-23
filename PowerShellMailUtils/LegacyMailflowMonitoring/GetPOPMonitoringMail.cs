using MailKit;
using MailKit.Net.Pop3;
using MailKit.Security;
using MimeKit;
using SyntheticTransactionsForExchange.DataModels;
using SyntheticTransactionsForExchange.Utilities;
using System;
using System.Diagnostics;
using System.IO;
using System.Management.Automation;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace SyntheticTransactionsForExchange.LegacyMailflowMonitoring
{
    [Cmdlet("Get", "POPMonitoringMail")]
    [OutputType(typeof(MailflowMonitoringData))]
    [CmdletBinding()]

    public class GetPOPMonitoringMail : Cmdlet
    {
        /// <summary>
        /// <para type="description">The FQDN or IP of the Server.</para>
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = @"The FQDN or IP of the Server")]
        public String Server;

        /// <summary>
        /// <para type="description">The port to which to connect.</para>
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = @"The port to which to connect")]
        public UInt16 Port;

        /// <summary>
        /// <para type="description">Username to access the Mailbox.</para>
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = @"Username to access the Mailbox")]
        public String UserName;

        /// <summary>
        /// <para type="description">Password to access the Mailbox.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"Password to access the Mailbox")]
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
        /// <para type="description">Guid to be set as Subject of the Message.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"Guid to be set as Subject of the Message.")]
        public Guid SubjectGuid;

        /// <summary>
        /// <para type="description">If specified the client will use SSL/TSL.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"If specified the client will use SSL/TSL")]
        public SwitchParameter UseTLS;

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
        /// <para type="description">If specified Protocol Logs would be collected.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"If specified tracing is enabled")]
        public SwitchParameter Trace;

        protected override void BeginProcessing()
        {
            WriteVerbose(String.Format("Entering cmdlet on <{0}>", DateTime.Now.ToString("yyyy-MM-MMTHH:mm:ss")));
        }

        protected override void ProcessRecord()
        {
            if (String.IsNullOrEmpty(Server))
            {
                throw new Exception("Server cannot be null");
            }

            if (Port == 0)
            {
                throw new Exception("Port cannot be zero");
            }
            if (Port != 110 && Port != 995)
            {
                WriteWarning(String.Format("POP usually works on port 110 (StartTLS) or 995 (POPS), is <{0}> correct?", Port));
            }

            if (String.IsNullOrEmpty(AccessToken) && Credentials == null && String.IsNullOrEmpty(Password))
            {
                throw new Exception("Authentication Missing, either specify AccessToken, Credentaials o Username/Password");
            }

            if (String.IsNullOrEmpty(UserName) && Credentials != null)
            {
                UserName = Credentials.UserName;
                WriteVerbose(String.Format("Setting UserName to match Credentials Username <{0}>", Credentials.UserName));
            }

            if (String.IsNullOrEmpty(UserName))
            {
                throw new Exception("Username cannot be null and must be provided");
            }

            if (SubjectGuid == null || SubjectGuid == Guid.Empty)
            {
                throw new Exception("Guid cannot be null");
            }

            MailflowMonitoringData monitoringData = new MailflowMonitoringData();

            DateTime receivedDateTime = DateTime.UtcNow;
            bool messageReceived = false;
            Stopwatch timer = new Stopwatch();
            MemoryStream logStream = new MemoryStream();

            try
            {
                using (Pop3Client client = new Pop3Client(new ProtocolLogger(logStream)))
                {
                    client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                    client.Connect(Server, Port, UseTLS);
                    if (!String.IsNullOrEmpty(AccessToken) && !String.IsNullOrEmpty(UserName))
                    {
                        client.Authenticate(new SaslMechanismOAuth2(UserName, AccessToken));
                    }
                    else
                    {
                        if (!String.IsNullOrEmpty(UserName) && !String.IsNullOrEmpty(Password))
                        {
                            client.Authenticate(UserName, Password);
                        }
                    }

                    timer.Start();

                    while (messageReceived == false && timer.ElapsedMilliseconds <= TimeOut * 1000)
                    {
                        for (int currentMessage = 0; currentMessage < client.Count && messageReceived == false; currentMessage++)
                        {
                            Stopwatch operationTimer = new Stopwatch();
                            operationTimer.Start();
                            MimeMessage message = client.GetMessage(currentMessage);
                            WriteVerbose(String.Format("Parsing message <{0}> sent from <{1}>", message.Subject, message.From));
                            operationTimer.Stop();
                            if (message.Subject.Contains(SubjectGuid.ToString()))
                            {
                                messageReceived = true;
                                WriteVerbose(String.Format("Found message <{0}> sent from <{1}>", message.Subject, message.From));
                                receivedDateTime = message.Date.UtcDateTime;
                                monitoringData.SetDuration(Convert.ToUInt32(operationTimer.ElapsedMilliseconds));
                                WriteVerbose(String.Format("The message body is <{0}>", message.TextBody));
                                MailflowMonitoringData temp = JsonSerializer.Deserialize<MailflowMonitoringData>(message.TextBody);
                                monitoringData.SetSendingInformation(temp.SubjectGuid, temp.TimeSent, temp.SendingProtocol);
                                string RAWHeader = String.Join(System.Environment.NewLine, message.Headers);
                                WriteVerbose(String.Format("The header parsed is <{0}> ", RAWHeader));
                                monitoringData.MailflowHeaderDataTable = MailUtilities.ParseMailHeader(RAWHeader);
                                monitoringData.SetStatus(TransactionStatus.Success);
                                client.DeleteMessage(currentMessage);
                            }
                        }
                        if (!messageReceived)
                        {
                            WriteVerbose(String.Format("Sleeping <{0}> seconds", SleepTimer));
                            Thread.Sleep(SleepTimer * 1000);
                        }
                        if (!messageReceived)
                        {
                            WriteWarning("The message has not been found!");
                        }
                    }
                    client.Disconnect(true);

                    if (Trace)
                    {
                        WriteVerbose("=========================== PROTOCOL LOG ===========================");
                        WriteVerbose(Encoding.UTF8.GetString(logStream.ToArray()));
                        WriteVerbose("====================================================================");
                    }
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
                if (Trace)
                {
                    WriteVerbose("=========================== PROTOCOL LOG ===========================");
                    WriteVerbose(Encoding.UTF8.GetString(logStream.ToArray()));
                    WriteVerbose("====================================================================");
                }
                throw ex;
            }

            monitoringData.SetReceivingInformation(SubjectGuid, receivedDateTime, Protocol.POP);

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
