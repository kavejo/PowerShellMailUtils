using Azure.Identity;
using Microsoft.Graph;
using SyntheticTransactionsForExchange.DataModels;
using SyntheticTransactionsForExchange.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Management.Automation;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace SyntheticTransactionsForExchange.GraphMailflowMonitoring
{
    [Cmdlet("Send", "GRAPHMonitoringMail")]
    [OutputType(typeof(MailflowMonitoringData))]
    [CmdletBinding()]

    public class SendGRAPHMonitoringMail : Cmdlet
    {
        /// <summary>
        /// <para type="description">The Url of the server to target with the Exchange Web Service connection. If missing autodiscover will be used.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"Exchange Web Services URL")]
        public String Url;

        /// <summary>
        /// <para type="description">The mailbox to target with this EWS Service. This parameter is mandatory.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"Mailbox to access via Exchange Web Services")]
        public String Mailbox;

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
        /// <para type="description">The list of recipients to whom to send the message.</para>
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = @"The list of recipients to whom to send the message.")]
        public List<String> Recipients;

        protected override void BeginProcessing()
        {
            WriteVerbose(String.Format("Entering cmdlet on <{0}>", DateTime.Now.ToString("yyyy-MM-MMTHH:mm:ss")));
        }

        protected override void ProcessRecord()
        {
            if (!String.IsNullOrEmpty(AccessToken) && String.IsNullOrEmpty(Mailbox))
            {
                throw new Exception("Mailbox and AccessToken must be both provided");
            }

            if (!RegexUtilities.IsValidEmail(Mailbox))
            {
                throw new Exception("Mailbox address is invalid");
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
            DelegateAuthenticationProvider authProvider = null;
            ClientSecretCredential clientSecretCredential= null;
            GraphServiceClient graphServiceClient = null;

            try
            {
                WriteVerbose(String.Format("Using OAuth token <{0}>", AccessToken));

                JwtSecurityTokenHandler jwtHandler = new JwtSecurityTokenHandler();
                JwtSecurityToken jwtToken = jwtHandler.ReadJwtToken(AccessToken);
                
                if (Guid.TryParse(jwtToken.Subject, out var newGuid))
                {
                    WriteVerbose(String.Format("Subject is a GUID, therefore it's Application permissions. Subject: <{0}>", jwtToken.Subject));
                }
                else
                {
                    WriteVerbose(String.Format("Subject is NOT a GUID, therefore it's Delegate permissions. Subject: <{0}>", jwtToken.Subject));
                }

                authProvider = new DelegateAuthenticationProvider(async (request) =>
                {
                    request.Headers.Authorization =
                        new AuthenticationHeaderValue("Bearer", AccessToken);
                });

                graphServiceClient = new GraphServiceClient(authProvider);
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

            DateTime sentDateTime = DateTime.UtcNow;

            try
            {

                monitoringData.SetSendingInformation(SubjectGuid, sentDateTime, Protocol.GRAPH);

                Message message = new Message();
                message.Subject = SubjectGuid.ToString();
                string jsonBody = JsonSerializer.Serialize(monitoringData);
                WriteVerbose(String.Format("The Serialized body content is <{0}>", jsonBody));
                message.Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = jsonBody
                };
                message.ToRecipients = new List<Recipient>();
                foreach (String recipient in Recipients)
                {
                    message.ToRecipients.Append(new Recipient { EmailAddress = new EmailAddress { Address = recipient } });
                }

                Stopwatch operationTimer = new Stopwatch();
                operationTimer.Start();

                WriteVerbose(String.Format("Sending mail message"));

                graphServiceClient
                    .Users[Mailbox]
                    .SendMail(message, true)
                    .Request()
                    .PostAsync()
                    .Wait();

                WriteVerbose(String.Format("Mail message sent"));

                operationTimer.Stop();
                monitoringData.SetDuration(Convert.ToUInt32(operationTimer.ElapsedMilliseconds));
                monitoringData.SetStatus(TransactionStatus.Success);
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

            WriteObject(monitoringData);
            return;
        }

        protected override void EndProcessing()
        {
            WriteVerbose(String.Format("Exiting cmdlet on <{0}>", DateTime.Now.ToString("yyyy-MM-MMTHH:mm:ss")));
        }
    }
}
