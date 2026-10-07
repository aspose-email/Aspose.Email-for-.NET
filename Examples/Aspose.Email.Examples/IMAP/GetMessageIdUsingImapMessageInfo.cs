// Demonstrates how to find a message by its Message-ID.
//
// The Message-ID stays with a message wherever it is copied or moved, so it is a
// reliable way to find "the same" message again - in another folder, or after the
// folder's unique ids were reset. ListMessage returns the Message-ID of one message, and
// HasHeader searches a folder for it on the server.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class GetMessageIdUsingImapMessageInfo
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                var ids = client.ListMessages(ImapFolderInfo.InBox, ImapListFields.IdOnly, 1);
                if (ids.Count == 0)
                {
                    Console.WriteLine("The Inbox is empty.");
                    return;
                }

                var info = client.ListMessage(ids[0].UniqueId);
                Console.WriteLine($"Message uid {info.UniqueId}: {info.Subject}");
                Console.WriteLine($"  Message-ID: {info.MessageId}");

                if (string.IsNullOrEmpty(info.MessageId))
                    return;

                var builder = new ImapQueryBuilder();
                builder.HasHeader("Message-ID", info.MessageId);

                var found = client.ListMessages(builder.GetQuery());
                Console.WriteLine($"\nSearching the Inbox for that Message-ID finds {found.Count} message(s):");
                foreach (var match in found)
                    Console.WriteLine($"  uid {match.UniqueId}: {match.Subject}");
            }
        }
    }
}
