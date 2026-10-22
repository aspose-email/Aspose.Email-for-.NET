// Demonstrates deleting a private distribution list.
//
// DeleteDistributionList removes the list itself - its members are not affected, they
// are just addresses. deletePermanently decides between Deleted Items and removing it
// for good. The example creates a uniquely named list and deletes it.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class DeletePrivateDistributionList
    {
        public static void Run()
        {
            var listName = "Aspose list " + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                client.CreateDistributionList(new ExchangeDistributionList { DisplayName = listName },
                    new MailAddressCollection { "member1@example.com" });

                var list = client.ListDistributionLists().First(l => l.DisplayName == listName);
                Console.WriteLine($"Created '{list.DisplayName}'.");

                client.DeleteDistributionList(list, true);

                var stillThere = client.ListDistributionLists().Any(l => l.DisplayName == listName);
                Console.WriteLine($"Deleted it; still present: {stillThere}");
            }
        }
    }
}
