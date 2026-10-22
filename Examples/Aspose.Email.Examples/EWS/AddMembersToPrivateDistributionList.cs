// Demonstrates adding members to an existing private distribution list.
//
// AddToDistributionList adds the given addresses to the list on the server. The example
// creates a uniquely named list, adds two members, shows the result and deletes the
// list at the end.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class AddMembersToPrivateDistributionList
    {
        public static void Run()
        {
            var listName = "Aspose list " + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                client.CreateDistributionList(new ExchangeDistributionList { DisplayName = listName },
                    new MailAddressCollection { "member1@example.com" });
                var list = client.ListDistributionLists().First(l => l.DisplayName == listName);

                try
                {
                    client.AddToDistributionList(list, new MailAddressCollection
                    {
                        "member2@example.com",
                        "member3@example.com"
                    });

                    Console.WriteLine($"Members of '{listName}' now:");
                    foreach (var member in client.FetchDistributionList(list))
                        Console.WriteLine($"  {member.Address}");
                }
                finally
                {
                    // The change key is outdated after the update; the id alone is enough.
                    client.DeleteDistributionList(new ExchangeDistributionList { Id = list.Id }, true);
                }
            }
        }
    }
}
