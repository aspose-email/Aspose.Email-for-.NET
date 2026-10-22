// Demonstrates printing the folder tree of a mailbox.
//
// ListSubFolders returns the direct children of a folder, so a recursive walk from the
// root reaches every level. Each ExchangeFolderInfo carries the display name, item and
// unread counts and the number of child folders, which saves a request for leaves.

using System;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class ListFoldersFromExchangeServer
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                Console.WriteLine("Folders of the mailbox:");
                foreach (var folder in client.ListSubFolders(client.MailboxInfo.RootUri))
                    Print(client, folder, 1);
            }
        }

        private static void Print(IEWSClient client, ExchangeFolderInfo folder, int depth)
        {
            Console.WriteLine($"{new string(' ', depth * 2)}{folder.DisplayName} " +
                              $"({folder.TotalCount} item(s), {folder.UnreadCount} unread)");

            if (folder.ChildFolderCount == 0)
                return;

            foreach (var child in client.ListSubFolders(folder.Uri))
                Print(client, child, depth + 1);
        }
    }
}
