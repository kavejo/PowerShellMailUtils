using SyntheticTransactionsForExchange.DataModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SyntheticTransactionsForExchange.Utilities
{
    public class MailUtilities
    {
        public static IList<MailflowHeaderData> ParseMailHeader(string MailHeader)
        {

            IList<MailflowHeaderData> headerDataList = new List<MailflowHeaderData>();

            Regex rx = new Regex(@"Received.(?:from) ?([\s\S]*?)?by([\s\S]*?)with([\s\S]*?)(?:;)([(\s\S)*]{32,36})(?:\s\S*?)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

            MatchCollection matches = rx.Matches(MailHeader);
            int Hop = 1;

            foreach (Match match in matches.Cast<Match>().Reverse())
            {
                GroupCollection groups = match.Groups;

                string regSubmittingHost = groups[1].Value;
                if (String.IsNullOrEmpty(regSubmittingHost))
                {
                    regSubmittingHost = groups[2].Value.Substring(0, groups[2].Value.IndexOf("by "));
                }

                string regReceivingHost = groups[2].Value;
                if (regReceivingHost.Contains("by"))
                {
                    regReceivingHost = groups[2].Value.Substring(groups[2].Value.IndexOf("by ") + 3);
                }

                string regType = groups[3].Value;

                string regReceivedTime = groups[4].Value.Replace(System.Environment.NewLine, "").Replace(" + ", "+").Trim();
                DateTime ReceivedTime = new DateTime();
                bool DateIsValid = DateTime.TryParse(regReceivedTime, out ReceivedTime);

                MailflowHeaderData headerData = new MailflowHeaderData();
                headerData.SubmittingHost = RegexUtilities.RemoveBetween(regSubmittingHost, '(', ')');
                headerData.DelayTime = TimeSpan.FromSeconds(0);
                headerData.ReceivingHost = RegexUtilities.RemoveBetween(regReceivingHost, '(', ')');
                headerData.Type = RegexUtilities.RemoveBetween(regType, '(', ')');
                headerData.Hop = Hop++;
                headerData.ReceivedTime = ReceivedTime;
                headerDataList.Add(headerData);

            }

            DateTime prev_item_ReceivedTime = new DateTime();
            foreach (var item in headerDataList)
            {
                if (item.ReceivedTime != new DateTime() && prev_item_ReceivedTime != new DateTime())
                {
                    item.DelayTime = item.ReceivedTime - prev_item_ReceivedTime;
                }
                prev_item_ReceivedTime = item.ReceivedTime;
            }

            return headerDataList;
        }
    }
}
