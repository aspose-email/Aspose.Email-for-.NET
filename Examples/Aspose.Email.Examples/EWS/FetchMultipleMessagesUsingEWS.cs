// Demonstrates downloading several messages in one request.
//
// FetchMessages takes message infos from a listing (or item URIs) and returns the
// complete messages - body and attachments - as MailMessage objects, saving a round trip
// per message compared with calling FetchMessage in a loop.

using System;
using System.Linq;

namespace Aspose.Email.Examples.EWS
{
    internal static class FetchMultipleMessagesUsingEWS
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var infos = client.ListMessages(client.MailboxInfo.InboxUri, 5);

                var messages = client.FetchMessages(infos.ToList());

                Console.WriteLine($"Fetched {messages.Count} message(s) from the Inbox:");
                foreach (var message in messages)
                {
                    Console.WriteLine($"  {message.Date:g}  {message.Subject}");
                    Console.WriteLine($"    from {message.From}, {message.Attachments.Count} attachment(s)");
                }
            }
        }
    }
}
