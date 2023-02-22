using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;
using System.Text.Json.Serialization;

namespace PowerShellMailUtils.DataModels
{
    internal class MailMessage
    {
        [JsonInclude]
        public string From { get; set; }
        [JsonInclude]
        public IList<string> To { get; set; }
        [JsonInclude]
        public IList<string> Cc { get; set; }
        [JsonInclude]
        public IList<string> Bcc { get; set; }
        [JsonInclude]
        public string Subject { get; set; }
        [JsonInclude]
        public string Body { get; set; }
        [JsonInclude]
        public bool IsValid { get; set; }

        public MailMessage(string Sender, IList<string> ToRecipients, string Subject, IList<string> CcRecipients = null, IList<string> BccRecipients = null, string Body = null)
        {
            this.From = String.Empty;
            this.To = new List<string>();
            this.Cc = new List<string>();
            this.Bcc = new List<string>();
            this.Subject = String.Empty;
            this.Body = String.Empty;

            if (ValidMailAddress(Sender))
                this.From = Sender;

            foreach (string address in ToRecipients)
                if (ValidMailAddress(address))
                    this.To.Add(address);

            foreach (string address in CcRecipients)
                if (ValidMailAddress(address))
                    this.Cc.Add(address);

            foreach (string address in BccRecipients)
                if (ValidMailAddress(address))
                    this.Bcc.Add(address);

            this.Subject = Subject;

            this.Body = Body;

        }

        public static bool ValidMailAddress(string address)
        {
            EmailAddressAttribute eMailVerify = new EmailAddressAttribute();
            return (eMailVerify.IsValid(address));
        }

        public bool ValidMessage()
        {
            IsValid = !String.IsNullOrEmpty(From) && To != null && To.Count >= 1 && !String.IsNullOrEmpty(Subject);
            return IsValid;
        }

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();

            sb.AppendLine(String.Format("From: <{0}>", From));

            foreach (string address in To)
                sb.AppendLine(String.Format("To: <{0}>", address));

            foreach (string address in Cc)
                sb.AppendLine(String.Format("Cc: <{0}>", address));

            foreach (string address in Bcc)
                sb.AppendLine(String.Format("Bcc: <{0}>", address));

            sb.AppendLine(String.Format("Subject: <{0}>", Subject));
            sb.AppendLine(String.Format("Body: <{0}>", Body));

            return sb.ToString();
        }
    }
}
