using System;
using System.Text;
using System.Text.Json.Serialization;

namespace SyntheticTransactionsForExchange.DataModels
{
    public class MailflowHeaderData
    {
        [JsonInclude]
        public int Hop { get; set; }
        [JsonInclude]
        public string SubmittingHost { get; set; }
        [JsonInclude]
        public string ReceivingHost { get; set; }
        [JsonInclude]
        public string Type { get; set; }
        [JsonInclude]
        public DateTime ReceivedTime { get; set; }
        [JsonInclude]
        public TimeSpan DelayTime { get; set; }

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(String.Format("Hop <{0}> from <{1}> to <{2}> via <{3}> on <{4}> took <{5}> seconds.", this.Hop, this.SubmittingHost, this.ReceivingHost, this.Type, this.ReceivedTime.ToString("yyyy-MM-MMTHH:mm:ss"), this.DelayTime.TotalSeconds));
            return sb.ToString();
        }

    }

}
