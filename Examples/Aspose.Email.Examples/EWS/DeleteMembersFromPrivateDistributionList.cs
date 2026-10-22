// Demonstrates removing members from a private distribution list.
//
// DeleteFromDistributionList takes member addresses as returned by
// FetchDistributionList - they carry the member ids Exchange needs to find them. The
// example creates a uniquely named list with three members, removes two and deletes
// the list at the end.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class DeleteMembersFromPrivateDistributionList
    {
        public static void Run()
        {
            var listName = "Aspose list " + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                client.CreateDistributionList(new ExchangeDistributionList { DisplayName = listName },
                    new MailAddressCollection { "member1@example.com", "member2@example.com", "member3@example.com" });
                var list = client.ListDistributionLists().First(l => l.DisplayName == listName);

                try
                {
                    var members = client.FetchDistributionList(list);
                    var toRemove = new MailAddressCollection { members[0], members[1] };

                    client.DeleteFromDistributionList(list, toRemove);

                    Console.WriteLine($"Removed {toRemove}; members of '{listName}' now:");
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
