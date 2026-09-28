// Demonstrates the lightest way to enumerate a folder.
//
// ImapListFields.IdOnly asks the server for nothing but sequence numbers and unique ids,
// skipping the envelope, flags and size that a normal listing downloads for every
// message. Use it when ids are all you need - for example to compare with a local
// cache - and fetch details only for the messages that matter.

using System;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ListMessageIdsOnly
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);
                var ids = client.ListMessages(ImapFolderInfo.InBox, ImapListFields.IdOnly, 100);
                Console.WriteLine($"{ids.Count} id(s) listed.");

                foreach (var info in ids.Take(10))
                    Console.WriteLine($"  #{info.SequenceNumber,-5} uid {info.UniqueId}");

                if (ids.Count == 0)
                    return;

                // Full summary information, for one message only.
                var details = client.ListMessage(ids[0].UniqueId);
                Console.WriteLine($"\nDetails of uid {details.UniqueId}:");
                Console.WriteLine($"  subject: {details.Subject}");
                Console.WriteLine($"  from:    {details.From}");
                Console.WriteLine($"  date:    {details.Date}");
                Console.WriteLine($"  size:    {details.Size} bytes");
            }
        }
    }
}
