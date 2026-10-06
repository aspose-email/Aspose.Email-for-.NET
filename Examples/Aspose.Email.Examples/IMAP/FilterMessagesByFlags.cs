// Demonstrates searching by message flags on the server.
//
// HasFlags adds a condition that the flags are set, HasNoFlags one that they are not;
// all conditions added to one builder must hold at the same time. Custom keywords set
// with ImapMessageFlags.Keyword can be searched for the same way.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class FilterMessagesByFlags
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                // Flagged for follow-up, but not answered yet.
                var builder = new ImapQueryBuilder();
                builder.HasFlags(ImapMessageFlags.Flagged);
                builder.HasNoFlags(ImapMessageFlags.Answered);

                var pending = client.ListMessages(builder.GetQuery());
                Console.WriteLine($"Flagged and unanswered: {pending.Count}");
                foreach (var info in pending)
                    Console.WriteLine($"  {info.Date:d}  {info.Subject}");

                // Unread and tagged with a custom keyword.
                builder = new ImapQueryBuilder();
                builder.HasFlags(ImapMessageFlags.Keyword("todo"));
                builder.HasNoFlags(ImapMessageFlags.IsRead);

                var todo = client.ListMessages(builder.GetQuery());
                Console.WriteLine($"\nUnread with keyword 'todo': {todo.Count}");
                foreach (var info in todo)
                    Console.WriteLine($"  {info.Date:d}  {info.Subject}");
            }
        }
    }
}
