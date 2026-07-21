using System;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class FetchPrivateDistributionList
    {
        public static void Run()
        {
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");
            ExchangeDistributionList[] distributionLists = client.ListDistributionLists();
            foreach (ExchangeDistributionList distributionList in distributionLists)
            {
                MailAddressCollection members = client.FetchDistributionList(distributionList);
                foreach (MailAddress member in members)
                {
                    Console.WriteLine(member.Address);
                }
            }
        }
    }
}