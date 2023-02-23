using MailKit;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using SyntheticTransactionsForExchange.DataModels;
using SyntheticTransactionsForExchange.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Text.Json;

namespace SyntheticTransactionsForExchange.LegacyMailflowMonitoring
{
    [Cmdlet("Send", "SMTPMonitoringMail")]
    [OutputType(typeof(MailflowMonitoringData))]
    [CmdletBinding()]

    public class SendSMTPMonitoringMail : Cmdlet
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
        [Parameter(Mandatory = false, HelpMessage = @"Username to access the Mailbox")]
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
        /// <para type="description">The sender address, in case it differs from the Username.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"The sender address, in case it differs from the Username")]
        public String SenderAddress;

        /// <summary>
        /// <para type="description">Guid to be set as Subject of the Message.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"Guid to be set as Subject of the Message.")]
        public Guid SubjectGuid;

        /// <summary>
        /// <para type="description">The list of recipients to whom to send the message.</para>
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = @"The list of recipients to whom to send the message.")]
        public List<String> Recipients;

        /// <summary>
        /// <para type="description">If specified the client will use SSL/TSL.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"If specified the client will use SSL/TSL")]
        public SwitchParameter UseTLS;

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
            if (Port != 25 && Port != 465 && Port != 587)
            {
                WriteWarning(String.Format("SMTP usually works on port 25 (StartTLS) or 465/587 (SMTPS), is <{0}> correct?", Port));
            }

            if (!String.IsNullOrEmpty(AccessToken) && String.IsNullOrEmpty(SenderAddress))
            {
                throw new Exception("SenderAddress cannot be null if using OAuth");
            }
            if (String.IsNullOrEmpty(SenderAddress) && !String.IsNullOrEmpty(UserName))
            {
                SenderAddress = UserName;
                WriteVerbose(String.Format("Setting SenderAddress to match Username <{0}>", UserName));
            }
            if (String.IsNullOrEmpty(SenderAddress) && Credentials != null)
            {
                SenderAddress = Credentials.UserName;
                WriteVerbose(String.Format("Setting SenderAddress to match Credentials Username <{0}>", Credentials.UserName));
            }
            if (String.IsNullOrEmpty(SenderAddress) && String.IsNullOrEmpty(AccessToken) && Credentials == null && String.IsNullOrEmpty(UserName))
            {
                SenderAddress = UserPrincipal.Current.EmailAddress;
                WriteVerbose(String.Format("Setting SenderAddress to the SMTP of the logged on user which is <{0}>", UserPrincipal.Current.EmailAddress));
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

            if (!RegexUtilities.IsValidEmail(SenderAddress))
            {
                throw new Exception("SenderAddress is not a valid SMTP address");
            }

            if (SubjectGuid == null || SubjectGuid == Guid.Empty)
            {
                SubjectGuid = Guid.NewGuid();
                WriteVerbose(String.Format("Guid Generated is <{0}>", SubjectGuid.ToString()));
            }

            if (Recipients == null)
            {
                throw new Exception("Recipients cannot be null");
            }
            if (Recipients.Any() == false)
            {
                throw new Exception("Recipients cannot be empty");
            }
            foreach (string recipient in Recipients)
            {
                if (!RegexUtilities.IsValidEmail(recipient))
                {
                    throw new Exception(String.Format("The recipient <{0}> is not a valid SMTP address", recipient));
                }
            }

            MailflowMonitoringData monitoringData = new MailflowMonitoringData();
            DateTime sentDateTime = DateTime.UtcNow;
            MemoryStream logStream = new MemoryStream();

            try
            {
                MimeMessage message = new MimeMessage();
                message.From.Add(MailboxAddress.Parse(SenderAddress));
                message.Subject = SubjectGuid.ToString();
                monitoringData.SetSendingInformation(SubjectGuid, sentDateTime, Protocol.SMTP);
                string jsonBody = JsonSerializer.Serialize(monitoringData);
                WriteVerbose(String.Format("The Serialized body content is <{0}>", jsonBody));
                message.Body = new TextPart("plain") { Text = jsonBody };

                foreach (String recipient in Recipients)
                {
                    message.To.Add(MailboxAddress.Parse(recipient));
                }

                using (SmtpClient client = new SmtpClient(new ProtocolLogger(logStream)))
                {
                    client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                    client.Connect(Server, Port, UseTLS);
                    if (!String.IsNullOrEmpty(AccessToken) && !String.IsNullOrEmpty(SenderAddress))
                    {
                        client.Authenticate(new SaslMechanismOAuth2(SenderAddress, AccessToken));
                    }
                    else
                    {
                        if (!String.IsNullOrEmpty(UserName) && !String.IsNullOrEmpty(Password))
                        {
                            client.Authenticate(UserName, Password);
                        }
                    }
                    Stopwatch operationTimer = new Stopwatch();
                    operationTimer.Start();
                    client.Send(message);
                    operationTimer.Stop();
                    monitoringData.SetDuration(Convert.ToUInt32(operationTimer.ElapsedMilliseconds));
                    monitoringData.SetStatus(TransactionStatus.Success);
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

            WriteObject(monitoringData);
            return;
        }

        protected override void EndProcessing()
        {
            WriteVerbose(String.Format("Exiting cmdlet on <{0}>", DateTime.Now.ToString("yyyy-MM-MMTHH:mm:ss")));
        }
    }
}
