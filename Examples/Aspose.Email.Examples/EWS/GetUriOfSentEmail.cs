// Demonstrates finding the copy of a sent message in Sent Items.
//
// Send does not return anything, but the ItemSent event reports the id of the copy that
// Exchange saved in Sent Items. Keep it to fetch, move or categorise the sent message
// later. The message goes to the configured mailbox itself.

using System;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class GetUriOfSentEmail
    {
        public static void Run()
        {
            var self = ClientBuilder.EwsUserName;
            string sentItemId = null;

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                client.ItemSent += (sender, e) => sentItemId = e.SentFolderItemId;

                client.Send(new MailMessage(self, self, "Track me", "Find this message in Sent Items."));

                if (sentItemId == null)
                {
                    Console.WriteLine("The ItemSent event did not report an id.");
                    return;
                }

                Console.WriteLine($"Sent Items id: {sentItemId}");
                Console.WriteLine($"Fetched back:  {client.FetchMessage(sentItemId).Subject}");
            }
        }
    }
}
