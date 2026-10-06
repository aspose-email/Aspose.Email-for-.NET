// Demonstrates how to get summary information for a known set of messages instead of a
// whole folder - for instance the ones a search or a local cache pointed to.
//
// Messages can be addressed by sequence number (their current position in the folder,
// which shifts when messages are removed) or by unique id (stable for as long as the
// folder's UIDVALIDITY does not change).

using System;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ListSpecificMessages
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);
                Console.WriteLine($"UIDVALIDITY of the Inbox: {client.CurrentFolder.ValidityId}");

                var ids = client.ListMessages(ImapFolderInfo.InBox, ImapListFields.IdOnly, 3);
                if (ids.Count == 0)
                {
                    Console.WriteLine("The Inbox is empty.");
                    return;
                }

                var sequenceNumbers = ids.Select(info => info.SequenceNumber).ToList();
                Console.WriteLine($"\nBy sequence number ({string.Join(", ", sequenceNumbers)}):");
                foreach (var info in client.ListMessages(ImapFolderInfo.InBox, sequenceNumbers))
                    Console.WriteLine($"  #{info.SequenceNumber}: {info.Subject}");

                var uniqueIds = ids.Select(info => info.UniqueId).ToList();
                Console.WriteLine($"\nBy unique id ({string.Join(", ", uniqueIds)}):");
                foreach (var info in client.ListMessages(ImapFolderInfo.InBox, uniqueIds))
                    Console.WriteLine($"  uid {info.UniqueId}: {info.Subject}");

                var single = client.ListMessage(sequenceNumbers[0]);
                Console.WriteLine($"\nOne message by sequence number: {single.Subject} ({single.Size} bytes)");
            }
        }
    }
}
