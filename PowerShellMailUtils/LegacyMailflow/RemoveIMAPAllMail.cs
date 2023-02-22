using MailKit;
using MailKit.Net.Imap;
using MailKit.Security;
using MimeKit;
using SyntheticTransactionsForExchange.DataModels;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Text;

namespace SyntheticTransactionsForExchange.LegacyMailflow
{
    [Cmdlet("Remove", "IMAPAllMail")]
    [OutputType(typeof(MailflowMonitoringData))]
    [CmdletBinding()]

    public class RemoveIMAPAllMail : Cmdlet
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
            if (Port != 143 && Port != 993)
            {
                WriteWarning(String.Format("IMAP usually works on port 143 (StartTLS) or 993 (IMAPS), is <{0}> correct?", Port));
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

            if (!String.IsNullOrEmpty(UserName))
            {
                throw new Exception("Username cannot be null and must be provided");
            }

            PerformanceMonitoringData monitoringData = new PerformanceMonitoringData();
            Stopwatch timer = new Stopwatch();
            MemoryStream logStream = new MemoryStream();

            try
            {
                using (ImapClient client = new ImapClient(new ProtocolLogger(logStream)))
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

                    IMailFolder inbox = client.Inbox;
                    inbox.Open(FolderAccess.ReadWrite);

                    timer.Start();
                    List<int> messagesList = new List<int>(Enumerable.Range(0, inbox.Count));
                    WriteVerbose(String.Format("Deleting <{0}> messages", inbox.Count));
                    client.Inbox.AddFlags(messagesList, MessageFlags.Deleted, true);
                    client.Inbox.Expunge();
                    client.Disconnect(true);
                    timer.Stop();
                    monitoringData.SetDuration(Convert.ToUInt32(timer.ElapsedMilliseconds));
                    monitoringData.SetStatus(TransactionStatus.Success);

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
            WriteVerbose(monitoringData.ToString());
            return;
        }
        protected override void EndProcessing()
        {
            WriteVerbose(String.Format("Exiting cmdlet on <{0}>", DateTime.Now.ToString("yyyy-MM-MMTHH:mm:ss")));
        }
    }
}
