using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json.Serialization;

namespace SyntheticTransactionsForExchange.DataModels
{

    public class MailflowMonitoringData
    {
        [JsonInclude]
        public Guid SubjectGuid;
        [JsonInclude]
        public DateTime TimeSent;
        [JsonInclude]
        public Protocol SendingProtocol;
        [JsonInclude]
        public DateTime TimeReceived;
        [JsonInclude]
        public Protocol ReceivingProtocol;
        [JsonInclude]
        public UInt32 Latency;
        [JsonInclude]
        public UInt32 CmdletDuration;
        [JsonInclude]
        public TransactionStatus Status;
        [JsonInclude]
        public IList<MailflowHeaderData> MailflowHeaderDataTable;

        public MailflowMonitoringData()
        {
            this.SubjectGuid = Guid.Empty;
            this.TimeSent = new DateTime();
            this.SendingProtocol = Protocol.None;
            this.TimeReceived = new DateTime();
            this.ReceivingProtocol = Protocol.None;
            this.Latency = 0;
            this.CmdletDuration = 0;
            this.Status = TransactionStatus.Failure;
            this.MailflowHeaderDataTable = new List<MailflowHeaderData>();
        }

        public MailflowMonitoringData(Guid subjectGuid, DateTime timeSent, Protocol sendingProtocol, DateTime timeReceived, Protocol receivingProtocol)
        {
            this.SubjectGuid = subjectGuid;
            this.TimeSent = timeSent;
            this.SendingProtocol = sendingProtocol;
            this.TimeReceived = timeReceived;
            this.ReceivingProtocol = receivingProtocol;
            this.Latency = 0;
            this.CmdletDuration = 0;
            this.Status = TransactionStatus.Failure;
        }

        public MailflowMonitoringData(Guid subjectGuid, DateTime timeSent, Protocol sendingProtocol, DateTime timeReceived, Protocol receivingProtocol, UInt32 latency)
        {
            this.SubjectGuid = subjectGuid;
            this.TimeSent = timeSent;
            this.SendingProtocol = sendingProtocol;
            this.TimeReceived = timeReceived;
            this.ReceivingProtocol = receivingProtocol;
            this.Latency = latency;
            this.CmdletDuration = 0;
            this.Status = TransactionStatus.Failure;
        }

        public MailflowMonitoringData(Guid subjectGuid, DateTime timeSent, Protocol sendingProtocol, DateTime timeReceived, Protocol receivingProtocol, UInt32 latency, UInt32 cmdletDuration)
        {
            this.SubjectGuid = subjectGuid;
            this.TimeSent = timeSent;
            this.SendingProtocol = sendingProtocol;
            this.TimeReceived = timeReceived;
            this.ReceivingProtocol = receivingProtocol;
            this.Latency = latency;
            this.CmdletDuration = cmdletDuration;
            this.Status = TransactionStatus.Failure;
        }

        public MailflowMonitoringData(Guid subjectGuid, DateTime timeSent, Protocol sendingProtocol, DateTime timeReceived, Protocol receivingProtocol, UInt32 latency, UInt32 cmdletDuration, TransactionStatus status)
        {
            this.SubjectGuid = subjectGuid;
            this.TimeSent = timeSent;
            this.SendingProtocol = sendingProtocol;
            this.TimeReceived = timeReceived;
            this.ReceivingProtocol = receivingProtocol;
            this.Latency = latency;
            this.CmdletDuration = cmdletDuration;
            this.Status = status;
        }

        public void SetSendingInformation(Guid subjectGuid, DateTime timeSent, Protocol sendingProtocol)
        {
            this.SubjectGuid = subjectGuid;
            this.TimeSent = timeSent;
            this.SendingProtocol = sendingProtocol;
        }

        public void SetReceivingInformation(Guid subjectGuid, DateTime timeReceived, Protocol receivingProtocol)
        {
            this.SubjectGuid = subjectGuid;
            this.TimeReceived = timeReceived;
            this.ReceivingProtocol = receivingProtocol;
        }

        public void SetDuration(UInt32 cmdletDuration)
        {
            this.CmdletDuration = cmdletDuration;
        }

        public void SetStatus(TransactionStatus status)
        {
            this.Status = status;
        }

        public void SetStatus(Boolean success)
        {
            if (success)
            {
                this.Status = TransactionStatus.Success;
            }
            else
            {
                this.Status = TransactionStatus.Failure;
            }
        }

        public void ComputeLatency()
        {
            double timeGap = (this.TimeReceived - this.TimeSent).TotalMilliseconds;
            this.Latency = (UInt32)Math.Ceiling(Math.Abs(timeGap));
        }


        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(String.Format("Message <{0}> sent via <{1}> at <{2}> and received via <{3}> on <{4}> took <{5}> seconds. Cmldet took <{6}> msec and status was <{7}>.", this.SubjectGuid, this.SendingProtocol, this.TimeSent.ToString("yyyy-MM-MMTHH:mm:ss"), this.ReceivingProtocol, this.TimeReceived.ToString("yyyy-MM-MMTHH:mm:ss"), Math.Round((double)this.Latency / 1000, 0), this.CmdletDuration, this.Status));
            if (MailflowHeaderDataTable != null && MailflowHeaderDataTable.Count >= 1)
            {
                foreach (MailflowHeaderData MailflowHeaderDataEntry in MailflowHeaderDataTable)
                    sb.Append(MailflowHeaderDataEntry.ToString());
            }
            return sb.ToString();
        }

    }
}
