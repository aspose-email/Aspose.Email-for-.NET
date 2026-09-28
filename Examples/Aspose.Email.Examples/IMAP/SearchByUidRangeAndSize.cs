// Demonstrates two IMAP-specific search keys: unique ids and message size.
//
// A UID range is the natural way to ask "what has arrived since I last looked":
// remember the highest UID you have seen and search from it up to '*' (the newest
// message). SimpleSeqSet picks individual UIDs. MessageSize finds messages larger or
// smaller than a number of bytes.

using System;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class SearchByUidRangeAndSize
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                var ids = client.ListMessages()
                    .OrderBy(info => long.Parse(info.UniqueId))
                    .ToList();

                if (ids.Count == 0)
                {
                    Console.WriteLine("The Inbox is empty.");
                    return;
                }

                // Pretend the last sync saw everything except the five newest messages.
                var lastSeenUid = ids[Math.Max(0, ids.Count - 6)].UniqueId;

                var builder = new ImapQueryBuilder();
                builder.UniqueId.Add(new RangeSeqSet(lastSeenUid, "*"));
                var newer = client.ListMessages(builder.GetQuery());

                Console.WriteLine($"UID {lastSeenUid}:* matches {newer.Count} message(s):");
                foreach (var info in newer)
                    Console.WriteLine($"  uid {info.UniqueId}: {info.Subject}");

                // Two individual UIDs.
                builder = new ImapQueryBuilder();
                builder.UniqueId.Add(new SimpleSeqSet(ids[0].UniqueId));
                builder.UniqueId.Add(new SimpleSeqSet(ids[ids.Count - 1].UniqueId));
                var oldestAndNewest = client.ListMessages(builder.GetQuery());

                Console.WriteLine("\nOldest and newest by UID:");
                foreach (var info in oldestAndNewest)
                    Console.WriteLine($"  uid {info.UniqueId}: {info.Subject}");

                // Everything over 5 MB.
                builder = new ImapQueryBuilder();
                builder.MessageSize.Greater(5 * 1024 * 1024);
                var large = client.ListMessages(builder.GetQuery());

                Console.WriteLine($"\nLarger than 5 MB: {large.Count} message(s)");
                foreach (var info in large)
                    Console.WriteLine($"  {info.Size / 1024 / 1024} MB  {info.Subject}");
            }
        }
    }
}
