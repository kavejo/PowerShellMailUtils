using Microsoft.Exchange.WebServices.Data;
using SyntheticTransactionsForExchange.DataModels;
using SyntheticTransactionsForExchange.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Management.Automation;

namespace SyntheticTransactionsForExchange.ExchangeCalendar
{
    [Cmdlet("Get", "EWSAvailability")]
    [OutputType(typeof(PerformanceMonitoringData))]
    [CmdletBinding()]

    public class GetEWSAvailability : Cmdlet
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
        /// <para type="description">The list of recipients to whom to send the message.</para>
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = @"The list of recipients for whom to check Free/Busy.")]
        public List<String> Recipients;

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

            PerformanceMonitoringData monitoringData = new PerformanceMonitoringData(DateTime.UtcNow);
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

            try
            {
                List<AttendeeInfo> attendees = new List<AttendeeInfo>();

                attendees.Add(new AttendeeInfo()
                {
                    SmtpAddress = Mailbox,
                    AttendeeType = MeetingAttendeeType.Organizer
                });

                foreach (string recipient in Recipients)
                {
                    attendees.Add(new AttendeeInfo()
                    {
                        SmtpAddress = recipient,
                        AttendeeType = MeetingAttendeeType.Organizer
                    });
                }

                AvailabilityOptions freeBusyOptions = new AvailabilityOptions();
                freeBusyOptions.MeetingDuration = 30;
                freeBusyOptions.RequestedFreeBusyView = FreeBusyViewType.FreeBusy;

                Stopwatch operationTimer = new Stopwatch();
                operationTimer.Start();
                GetUserAvailabilityResults freeBusyResults = EWSService.GetUserAvailability(attendees, new TimeWindow(DateTime.Now, DateTime.Now.AddDays(7)), AvailabilityData.FreeBusy, freeBusyOptions);
                operationTimer.Stop();

                monitoringData.SetDuration(Convert.ToUInt32(operationTimer.ElapsedMilliseconds));
                monitoringData.SetStatus(freeBusyResults != null && freeBusyResults.AttendeesAvailability.Count > 0 && freeBusyResults.AttendeesAvailability.OverallResult == ServiceResult.Success);

                if (freeBusyResults.AttendeesAvailability.OverallResult != ServiceResult.Success)
                {
                    WriteWarning(String.Format("Error has been thrown: <{0}>", freeBusyResults.AttendeesAvailability.OverallResult));
                    foreach (AttendeeAvailability availability in freeBusyResults.AttendeesAvailability)
                    {
                        WriteWarning(String.Format("The error message is: <{0}>", availability.ErrorMessage));
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
