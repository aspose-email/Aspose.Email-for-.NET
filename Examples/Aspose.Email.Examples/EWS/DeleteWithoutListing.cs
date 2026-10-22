// Demonstrates deleting a distribution list known only by its id.
//
// An ExchangeDistributionList with just the stored Id is enough for
// DeleteDistributionList, so there is no need to list all distribution lists first.
// The example creates a uniquely named list, "stores" its id, and deletes it.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class DeleteWithoutListing
    {
        public static void Run()
        {
            var listName = "Aspose list " + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                client.CreateDistributionList(new ExchangeDistributionList { DisplayName = listName },
                    new MailAddressCollection { "member1@example.com" });

                // Normally saved earlier, for example in your database.
                var storedId = client.ListDistributionLists().First(l => l.DisplayName == listName).Id;

                client.DeleteDistributionList(new ExchangeDistributionList { Id = storedId }, true);

                var stillThere = client.ListDistributionLists().Any(l => l.DisplayName == listName);
                Console.WriteLine($"Deleted '{listName}' by id; still present: {stillThere}");
            }
        }
    }
}
