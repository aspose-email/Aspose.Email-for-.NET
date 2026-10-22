// Demonstrates moving a message to another folder.
//
// MoveItem takes the item URI first and the destination folder second, and returns the
// item's new URI - an item gets a new id when it changes folders, so keep using the
// returned one. The example works in two uniquely named folders and deletes both at the
// end.

using System;
using System.Collections.Specialized;

namespace Aspose.Email.Examples.EWS
{
    internal static class MoveMessageFromOneFolderToAnotherusingEWS
    {
        public static void Run()
        {
            var suffix = Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var source = client.CreateFolder(client.MailboxInfo.InboxUri, "Aspose-Inbox-" + suffix);
                var processed = client.CreateFolder(client.MailboxInfo.InboxUri, "Aspose-Processed-" + suffix);

                try
                {
                    var uri = client.AppendMessage(source.Uri,
                        new MailMessage("from@example.com", "to@example.com", "Process this message", "Body"));

                    var newUri = client.MoveItem(uri, processed.Uri);
                    Console.WriteLine($"Moved: {client.FetchMessage(newUri).Subject}");

                    Console.WriteLine($"  '{source.DisplayName}': {client.ListMessages(source.Uri).Count} message(s)");
                    var processedCount = client.ListMessages(processed.Uri).Count;
                    Console.WriteLine($"  '{processed.DisplayName}': {processedCount} message(s)");
                }
                finally
                {
                    client.DeleteFolders(new StringCollection { source.Uri, processed.Uri }, true);
                }
            }
        }
    }
}
