using System;
using System.Text.Json.Serialization;

namespace SyntheticTransactionsForExchange.DataModels
{

    public class PerformanceMonitoringData
    {
        [JsonInclude]
        public DateTime DateTime;
        [JsonInclude]
        public UInt32 CmdletDuration;
        [JsonInclude]
        public TransactionStatus Status;

        public PerformanceMonitoringData()
        {
            this.DateTime = DateTime.UtcNow;
            this.CmdletDuration = 0;
            this.Status = TransactionStatus.Failure;
        }

        public PerformanceMonitoringData(DateTime dateTime)
        {
            this.DateTime = dateTime;
            this.CmdletDuration = 0;
            this.Status = TransactionStatus.Failure;
        }

        public PerformanceMonitoringData(DateTime dateTime, UInt32 cmdletDuration, TransactionStatus status)
        {
            this.DateTime = dateTime;
            this.CmdletDuration = cmdletDuration;
            this.Status = status;
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

        public override String ToString()
        {
            return String.Format("The operation executed on <{0}> has been executed in <{1}> msec and terminated with <{2}>.", this.DateTime.ToString("yyyy-MM-MMTHH:mm:ss"), this.CmdletDuration, this.Status);
        }

    }
}
