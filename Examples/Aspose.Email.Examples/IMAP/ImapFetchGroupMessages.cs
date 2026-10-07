// Demonstrates downloading several messages with one call.
//
// FetchMessages takes a list of sequence numbers or unique ids and returns the complete
// messages, which saves a round trip per message compared with calling FetchMessage in
// a loop. Sequence numbers are positions in the folder and change when messages are
// removed; unique ids stay stable.

using System;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ImapFetchGroupMessages
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                var newest = client.ListMessages()
                    .OrderByDescending(info => info.InternalDate)
                    .Take(5)
                    .ToList();

                var bySequenceNumber = client.FetchMessages(newest.Select(info => info.SequenceNumber).ToList());
                Console.WriteLine($"Fetched by sequence number: {bySequenceNumber.Count} message(s)");

                var byUniqueId = client.FetchMessages(newest.Select(info => info.UniqueId).ToList());
                Console.WriteLine($"Fetched by unique id:       {byUniqueId.Count} message(s)");

                foreach (var message in byUniqueId)
                    Console.WriteLine($"  {message.Date:g}  {message.Subject}");
            }
        }
    }
}
