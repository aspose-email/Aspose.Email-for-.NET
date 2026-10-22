// Demonstrates deleting folders, one at a time and several in one request.
//
// DeleteFolder moves the folder with its contents to Deleted Items, or removes it for
// good when deletePermanently is true. DeleteFolders does the same for a list of
// folders in a single EWS call. The example creates uniquely named folders under the
// Inbox and deletes them.

using System;
using System.Collections.Specialized;

namespace Aspose.Email.Examples.EWS
{
    internal static class DeleteFoldersOnExchange
    {
        public static void Run()
        {
            var suffix = Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var inboxUri = client.MailboxInfo.InboxUri;

                var single = client.CreateFolder(inboxUri, "Aspose-Single-" + suffix);
                var first = client.CreateFolder(inboxUri, "Aspose-Batch1-" + suffix);
                var second = client.CreateFolder(inboxUri, "Aspose-Batch2-" + suffix);
                Console.WriteLine($"Created {single.DisplayName}, {first.DisplayName} and {second.DisplayName}.");

                // To Deleted Items, where it can still be recovered.
                client.DeleteFolder(single.Uri, false);
                Console.WriteLine($"Moved '{single.DisplayName}' to Deleted Items.");

                // Two folders in one request, removed for good.
                client.DeleteFolders(new StringCollection { first.Uri, second.Uri }, true);
                Console.WriteLine($"Deleted '{first.DisplayName}' and '{second.DisplayName}' permanently.");

                var stillExists = client.FolderExists(inboxUri, first.DisplayName);
                Console.WriteLine($"\n'{first.DisplayName}' still exists: {stillExists}");
            }
        }
    }
}
