using System;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class DeleteMembersWithoutListing
    {
        public static void Run()
        {
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");
            ExchangeDistributionList distributionList = new ExchangeDistributionList();
            distributionList.Id = "list's id";
            distributionList.ChangeKey = "list's change key";
            MailAddressCollection membersToDelete = new MailAddressCollection();
            MailAddress addressToDelete = new MailAddress("address", true);
            //addressToDelete.Id.EWSId = "member's id";
            membersToDelete.Add(addressToDelete);
            client.AddToDistributionList(distributionList, membersToDelete);
        }
    }
}