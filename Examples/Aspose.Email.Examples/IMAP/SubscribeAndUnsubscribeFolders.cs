// Demonstrates folder subscriptions.
//
// The subscription list is the set of folders a user wants their mail programs to show.
// It is kept on the server, so every client the user runs sees the same list.
// SubscribeFolder and UnsubscribeFolder edit it; on servers with LIST-EXTENDED
// (RFC 5258), ListFolders with the Subscribed selection option reads it back.

using System;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class SubscribeAndUnsubscribeFolders
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.CreateFolder(folderName);

                try
                {
                    client.SubscribeFolder(folderName);
                    Console.WriteLine($"Subscribed to '{folderName}'.");
                    ReportSubscription(client, folderName);

                    client.UnsubscribeFolder(folderName);
                    Console.WriteLine($"Unsubscribed from '{folderName}'.");
                    ReportSubscription(client, folderName);
                }
                finally
                {
                    client.DeleteFolder(folderName);
                }
            }
        }

        private static void ReportSubscription(ImapClient client, string folderName)
        {
            if (!client.ExtendedListSupported)
                return;

            // A null parent folder lists from the top of the hierarchy.
            var subscribed = client.ListFolders(null, false,
                ListFoldersOptions.Subscribed, ListFoldersReturnOptions.None);
            var isListed = subscribed.Any(folder => folder.Name == folderName);

            Console.WriteLine($"  on the server's subscription list: {isListed}");
            Console.WriteLine($"  subscribed folders in total:       {subscribed.Count}");
        }
    }
}
