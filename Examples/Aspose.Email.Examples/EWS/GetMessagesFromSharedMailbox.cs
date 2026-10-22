// Demonstrates reading messages from a shared mailbox.
//
// The ListItems overload that takes a mailbox address lists items of a folder in that
// mailbox, and FetchItem downloads each one as a MapiMessage. The signed-in account
// needs at least read permission on the shared mailbox. Replace the address below.

using System;
using System.Linq;

namespace Aspose.Email.Examples.EWS
{
    internal static class GetMessagesFromSharedMailbox
    {
        public static void Run()
        {
            const string sharedMailbox = "shared@example.com";

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var items = client.ListItems(sharedMailbox, "Inbox");
                Console.WriteLine($"{items.Length} item(s) in the Inbox of {sharedMailbox}, the first 5:");

                foreach (var uri in items.Take(5))
                {
                    var message = client.FetchItem(uri);
                    Console.WriteLine($"  {message.SenderName}: {message.Subject}");
                }
            }
        }
    }
}
