using Aspose.Email.Clients.Exchange;
using System;

namespace Aspose.Email.Examples.EWS
{
    class AddMembersToPrivateDistributionList
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var distributionLists = client.ListDistributionLists();

                if (distributionLists.Length != 0)
                {
                    var newMembers = new MailAddressCollection
                    {
                        "address4@host.com",
                        "address5@host.com"
                    };
                    
                    client.AddToDistributionList(distributionLists[0], newMembers);
                }
                else
                {
                    Console.WriteLine(@"There is no distribution lists in ""Contacts"" folder");
                }
            }
        }
    }
}