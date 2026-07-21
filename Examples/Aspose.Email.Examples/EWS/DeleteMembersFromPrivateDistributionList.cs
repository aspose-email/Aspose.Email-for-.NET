using System;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class DeleteMembersFromPrivateDistributionList
    {
        public static void Run()
        {
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");
            ExchangeDistributionList[] distributionLists = client.ListDistributionLists();
            MailAddressCollection members = client.FetchDistributionList(distributionLists[0]);
            MailAddressCollection membersToDelete = new MailAddressCollection();
            membersToDelete.Add(members[0]);
            membersToDelete.Add(members[1]);
            client.DeleteFromDistributionList(distributionLists[0], membersToDelete);
        }
    }
}