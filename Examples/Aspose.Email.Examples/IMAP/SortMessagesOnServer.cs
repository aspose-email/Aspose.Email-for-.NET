// Demonstrates server-side sorting with the SORT extension (RFC 5256).
//
// The server returns the matching messages already in the requested order, so the
// client does not have to download every envelope just to sort. SortBy and ReverseBy
// take SortingKey values; a key given in ReverseBy sorts in descending order.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class SortMessagesOnServer
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                if (!client.SortSupported)
                {
                    Console.WriteLine("The server does not support SORT.");
                    return;
                }

                var bySubject = new SortConditions
                {
                    SortBy = SortingKey.Subject,
                    UseUId = true,
                    Since = DateTime.Today.AddDays(-30)
                };
                Print(client, "Last 30 days by subject, A to Z:", client.SortMessageThreads(bySubject));

                var largestFirst = new SortConditions
                {
                    ReverseBy = SortingKey.Size,
                    UseUId = true,
                    Since = DateTime.Today.AddDays(-30)
                };
                Print(client, "Last 30 days by size, largest first:", client.SortMessageThreads(largestFirst));
            }
        }

        private static void Print(ImapClient client, string title, List<MessageThreadResult> sorted)
        {
            var top = sorted.Take(10).Select(result => result.UniqueId).ToList();

            // One round trip for the details; the loop below keeps the server's order.
            var details = client.ListMessages(ImapFolderInfo.InBox, top).ToDictionary(info => info.UniqueId);

            Console.WriteLine($"{title} ({sorted.Count} message(s), first {top.Count} shown)");
            foreach (var uid in top)
            {
                ImapMessageInfo info;
                if (details.TryGetValue(uid, out info))
                    Console.WriteLine($"  {info.Size,10:N0} bytes  {info.Subject}");
            }
            Console.WriteLine();
        }
    }
}
