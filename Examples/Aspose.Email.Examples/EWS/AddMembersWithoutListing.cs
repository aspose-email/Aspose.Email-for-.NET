// Demonstrates adding members to a distribution list known only by its id.
//
// A distribution list is identified by Id and ChangeKey. Store both when you first see
// the list, and later an ExchangeDistributionList with just those two values is enough
// for AddToDistributionList - no need to list all distribution lists again. The example
// creates a uniquely named list, "stores" its ids, and deletes the list at the end.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class AddMembersWithoutListing
    {
        public static void Run()
        {
            var listName = "Aspose list " + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                client.CreateDistributionList(new ExchangeDistributionList { DisplayName = listName },
                    new MailAddressCollection { "member1@example.com" });

                // Normally saved earlier, for example in your database.
                var created = client.ListDistributionLists().First(l => l.DisplayName == listName);
                var storedId = created.Id;
                var storedChangeKey = created.ChangeKey;

                var list = new ExchangeDistributionList { Id = storedId, ChangeKey = storedChangeKey };

                try
                {
                    client.AddToDistributionList(list, new MailAddressCollection { "member2@example.com" });

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
