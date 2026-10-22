// Demonstrates creating nested folders with one call.
//
// With UseSlashAsFolderSeparator set, a name such as "Projects/2026/Q1" passed to
// CreateFolder creates every missing level of the path; without it the slash is just a
// character of a single folder name. The example creates the tree under a uniquely
// named folder and deletes it at the end.

using System;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class CreateFoldersOnExchangeServerMailbox
    {
        public static void Run()
        {
            var rootName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var root = client.CreateFolder(client.MailboxInfo.InboxUri, rootName);

                try
                {
                    client.UseSlashAsFolderSeparator = true;
                    client.CreateFolder(root.Uri, "Projects/2026/Q1");
                    client.CreateFolder(root.Uri, "Projects/2026/Q2");

                    Console.WriteLine($"Folder tree under '{rootName}':");
                    Print(client, root.Uri, 1);
                }
                finally
                {
                    client.DeleteFolder(root.Uri, true);
                }
            }
        }

        private static void Print(IEWSClient client, string folderUri, int depth)
        {
            foreach (var folder in client.ListSubFolders(folderUri))
            {
                Console.WriteLine(new string(' ', depth * 2) + folder.DisplayName);
                Print(client, folder.Uri, depth + 1);
            }
        }
    }
}
