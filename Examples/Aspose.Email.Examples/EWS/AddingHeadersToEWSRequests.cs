// Demonstrates adding an HTTP header to every EWS request.
//
// Exchange Online uses the X-AnchorMailbox header to route a request straight to the
// server that holds the target mailbox. Microsoft recommends sending it when working
// with another mailbox - impersonation, shared mailboxes - to avoid failed or slow
// requests. Replace the address below with the mailbox you work with.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class AddingHeadersToEWSRequests
    {
        public static void Run()
        {
            const string targetMailbox = "user@example.com";

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                client.AddHeader("X-AnchorMailbox", targetMailbox);

                var messages = client.ListMessages(client.MailboxInfo.InboxUri, 5);
                Console.WriteLine($"Listed {messages.Count} message(s) with X-AnchorMailbox: {targetMailbox}");
            }
        }
    }
}
