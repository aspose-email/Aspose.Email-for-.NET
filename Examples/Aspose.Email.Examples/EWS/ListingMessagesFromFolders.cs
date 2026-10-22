// Demonstrates listing messages from several well-known folders.
//
// MailboxInfo holds the URI of every well-known folder, so the same ListMessages call
// works for the Inbox, Sent Items, Drafts or Deleted Items. The second argument caps
// how many summaries are returned.

using System;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class ListingMessagesFromFolders
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var mailbox = client.MailboxInfo;

                Print(client, "Inbox", mailbox.InboxUri);
                Print(client, "Sent Items", mailbox.SentItemsUri);
                Print(client, "Drafts", mailbox.DraftsUri);
                Print(client, "Deleted Items", mailbox.DeletedItemsUri);
            }
        }

        private static void Print(IEWSClient client, string name, string folderUri)
        {
            var messages = client.ListMessages(folderUri, 3);

            Console.WriteLine($"{name}:");
            foreach (var info in messages)
                Console.WriteLine($"  {info.InternalDate:g}  {info.Subject}");
            if (messages.Count == 0)
                Console.WriteLine("  (empty)");
        }
    }
}
