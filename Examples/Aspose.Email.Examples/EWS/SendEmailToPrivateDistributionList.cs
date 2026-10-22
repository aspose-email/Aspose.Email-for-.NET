// Demonstrates sending a message to a private distribution list.
//
// ToMailAddress turns an ExchangeDistributionList into an address that can be put into
// To, Cc or Bcc; Exchange expands it to the members when sending. The example creates a
// uniquely named list whose only member is the configured mailbox, sends to it, and
// deletes the list at the end.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class SendEmailToPrivateDistributionList
    {
        public static void Run()
        {
            var self = ClientBuilder.EwsUserName;
            var listName = "Aspose list " + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                client.CreateDistributionList(new ExchangeDistributionList { DisplayName = listName },
                    new MailAddressCollection { self });
                var list = client.ListDistributionLists().First(l => l.DisplayName == listName);

                try
                {
                    var message = new MailMessage(new MailAddress(self), list.ToMailAddress())
                    {
                        Subject = "Message to a distribution list",
                        Body = "Everyone on the list receives this."
                    };

                    client.Send(message);
                    Console.WriteLine($"Sent to the list '{listName}'.");
                }
                finally
                {
                    client.DeleteDistributionList(list, true);
                }
            }
        }
    }
}
