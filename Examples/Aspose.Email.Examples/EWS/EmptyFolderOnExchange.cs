// Demonstrates emptying a folder without deleting the folder itself.
//
// EmptyFolder removes every item in the folder. EmptyFolderOptions decides whether
// subfolders go too (DeleteSubfolders) and whether the items are removed for good
// (DeletePermanently) or moved to Deleted Items. The example fills a uniquely named
// folder, empties it and deletes it at the end.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class EmptyFolderOnExchange
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.InboxUri, folderName);

                try
                {
                    client.CreateFolder(folder.Uri, "Subfolder");

                    var messages = Enumerable.Range(1, 3)
                        .Select(i => new MailMessage("sender@example.com", "receiver@example.com",
                            $"Old message {i}", "Body"));
                    client.AppendMessages(folder.Uri, messages);

                    Report(client, folder.Uri, "Before:");

                    var options = EmptyFolderOptions.DeleteSubfolders | EmptyFolderOptions.DeletePermanently;
                    client.EmptyFolder(folder.Uri, options);

                    Report(client, folder.Uri, "After: ");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }

        private static void Report(IEWSClient client, string folderUri, string title)
        {
            var info = client.GetFolderInfo(folderUri);
            Console.WriteLine($"{title} {info.TotalCount} item(s), {info.ChildFolderCount} subfolder(s)");
        }
    }
}
