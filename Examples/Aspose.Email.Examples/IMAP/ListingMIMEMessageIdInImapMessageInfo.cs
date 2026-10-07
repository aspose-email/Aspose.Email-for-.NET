// Demonstrates how to list the Message-ID of every message in a folder.
//
// The Message-ID header is a globally unique identifier the sending program gives a
// message. Unlike the IMAP unique id, which only identifies a message within one folder,
// it stays the same when the message is copied or moved - which makes it the key for
// spotting duplicates or matching replies. ImapMessageInfo.MessageId carries it.

using System;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ListingMIMEMessageIdInImapMessageInfo
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                var messages = client.ListMessages(ImapFolderInfo.InBox);

                Console.WriteLine($"{messages.Count} message(s) in the Inbox, the first 10:");
                foreach (var info in messages.Take(10))
                    Console.WriteLine($"  uid {info.UniqueId,-8} {info.MessageId}");

                var duplicates = messages
                    .Where(info => !string.IsNullOrEmpty(info.MessageId))
                    .GroupBy(info => info.MessageId)
                    .Count(group => group.Count() > 1);

                Console.WriteLine($"\nMessage-IDs that occur more than once: {duplicates}");
            }
        }
    }
}
