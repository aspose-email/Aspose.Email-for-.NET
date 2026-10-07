// Demonstrates listing a folder newest first.
//
// PageSettings.AscendingSorting decides the order in which ListMessagesByPage walks the
// folder: true starts with the oldest messages, false with the newest. With false the
// first page is what an inbox view shows, without listing the whole folder.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ChangeOrderOfEmails
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                var oldestFirst = new PageSettings { FolderName = ImapFolderInfo.InBox, AscendingSorting = true };
                var newestFirst = new PageSettings { FolderName = ImapFolderInfo.InBox, AscendingSorting = false };

                Print("Oldest first:", client.ListMessagesByPage(5, oldestFirst));
                Print("Newest first:", client.ListMessagesByPage(5, newestFirst));
            }
        }

        private static void Print(string title, ImapPageInfo page)
        {
            Console.WriteLine(title);
            foreach (var info in page.Items)
                Console.WriteLine($"  {info.Date:g}  {info.Subject}");
            Console.WriteLine();
        }
    }
}
