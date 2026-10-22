// Demonstrates measuring how much space a folder takes, including its subfolders.
//
// GetMailboxSizeEx adds up the size of the given folder and every folder below it, so
// passing the root gives the size of the whole mailbox. It walks the tree folder by
// folder and can take a while on large mailboxes. For the single-call GetMailboxSize see
// GetMailboxInformationFromExchangeWebServices.

using System;
using System.Linq;

namespace Aspose.Email.Examples.EWS
{
    internal static class GetFolderSizeRecursively
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var inbox = client.MailboxInfo.InboxUri;

                Console.WriteLine($"Inbox with subfolders: {ToMegabytes(client.GetMailboxSizeEx(inbox))}");

                foreach (var folder in client.ListSubFolders(inbox).Take(10))
                    Console.WriteLine($"  {folder.DisplayName}: {ToMegabytes(client.GetMailboxSizeEx(folder.Uri))}");

                var mailboxSize = client.GetMailboxSizeEx(client.MailboxInfo.RootUri);
                Console.WriteLine($"\nWhole mailbox: {ToMegabytes(mailboxSize)}");
            }
        }

        private static string ToMegabytes(long bytes)
        {
            return $"{bytes / 1024.0 / 1024.0:F1} MB";
        }
    }
}
