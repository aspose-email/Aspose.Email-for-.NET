// Demonstrates creating a private distribution list (contact group).
//
// A private distribution list lives in the user's Contacts folder, like an Outlook
// contact group; sending to it reaches all its members. CreateDistributionList takes
// the list and its members. The example creates a uniquely named list, reads its
// members back and deletes it at the end.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class CreatePrivateDistributionList
    {
        public static void Run()
        {
            var listName = "Aspose list " + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var members = new MailAddressCollection
                {
                    "member1@example.com",
                    "member2@example.com",
                    "member3@example.com"
                };
                client.CreateDistributionList(new ExchangeDistributionList { DisplayName = listName }, members);

                var list = client.ListDistributionLists().First(l => l.DisplayName == listName);

                try
                {
                    Console.WriteLine($"Created '{list.DisplayName}' with:");
                    foreach (var member in client.FetchDistributionList(list))
                        Console.WriteLine($"  {member.Address}");
                }
                finally
                {
                    client.DeleteDistributionList(list, true);
                }
            }
        }
    }
}
