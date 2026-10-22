// Demonstrates listing the messages of a folder.
//
// ListMessages returns a summary per message - subject, sender, recipients, date,
// Message-ID and the UniqueUri that identifies the item for fetching, moving or
// deleting - without downloading the messages themselves.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class ListingMessagesUsingEWS
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var messages = client.ListMessages(client.MailboxInfo.InboxUri, 10);
                Console.WriteLine($"{messages.Count} message(s) from the Inbox:");

                foreach (var info in messages)
                {
                    Console.WriteLine($"\n  {info.Subject}");
                    Console.WriteLine($"    from:       {info.From}");
                    Console.WriteLine($"    to:         {info.To}");
                    Console.WriteLine($"    received:   {info.InternalDate:g}");
                    Console.WriteLine($"    message id: {info.MessageId}");
                    Console.WriteLine($"    unique uri: {info.UniqueUri}");
                }
            }
        }
    }
}
