// Demonstrates removing members from a distribution list known only by its id.
//
// With the list's stored Id and ChangeKey there is no need to list all distribution
// lists first. The members to remove still come from FetchDistributionList, because
// each returned address carries the member id Exchange matches on. The example creates
// a uniquely named list, "stores" its ids, and deletes the list at the end.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class DeleteMembersWithoutListing
    {
        public static void Run()
        {
            var listName = "Aspose list " + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                client.CreateDistributionList(new ExchangeDistributionList { DisplayName = listName },
                    new MailAddressCollection { "member1@example.com", "member2@example.com" });

                // Normally saved earlier, for example in your database.
                var created = client.ListDistributionLists().First(l => l.DisplayName == listName);
                var list = new ExchangeDistributionList { Id = created.Id, ChangeKey = created.ChangeKey };

                try
                {
                    var toRemove = new MailAddressCollection();
                    toRemove.Add(client.FetchDistributionList(list).First(m => m.Address == "member1@example.com"));

                    client.DeleteFromDistributionList(list, toRemove);

                    Console.WriteLine($"Members of '{listName}' after removing member1:");
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
