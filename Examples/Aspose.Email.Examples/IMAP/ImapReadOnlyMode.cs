// Demonstrates reading messages without changing anything on the server.
//
// Fetching a message normally sets its \Seen flag, so it stops showing as unread in the
// user's mail program. With ReadOnly set, the client selects folders read-only
// (EXAMINE), and the server leaves every flag as it was - which suits archiving,
// indexing or monitoring tools.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ImapReadOnlyMode
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.ReadOnly = true;
                client.SelectFolder(ImapFolderInfo.InBox);

                var builder = new ImapQueryBuilder();
                builder.HasNoFlags(ImapMessageFlags.IsRead);
                var unreadQuery = builder.GetQuery();

                var unread = client.ListMessages(unreadQuery);
                Console.WriteLine($"Unread before fetching: {unread.Count}");

                if (unread.Count == 0)
                {
                    Console.WriteLine("No unread messages to try it on.");
                    return;
                }

                var message = client.FetchMessage(unread[0].UniqueId);
                Console.WriteLine($"Fetched: {message.Subject}");

                Console.WriteLine($"Unread after fetching:  {client.ListMessages(unreadQuery).Count}");
            }
        }
    }
}
