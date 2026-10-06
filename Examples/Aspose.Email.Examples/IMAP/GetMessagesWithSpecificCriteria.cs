// Demonstrates common search criteria: arrival date, sender, recipient and flags.
//
// Each query uses a fresh builder, and all conditions added to one builder must hold at
// once. MailQueryBuilder holds the criteria every mail protocol understands;
// ImapQueryBuilder derives from it and adds IMAP-only ones such as HasFlags and
// HasNoFlags.

using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Tools.Search;

namespace Aspose.Email.Examples.IMAP
{
    internal static class GetMessagesWithSpecificCriteria
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                var builder = new MailQueryBuilder();
                builder.InternalDate.On(DateTime.Now);
                Search(client, "Arrived today", builder.GetQuery());

                builder = new MailQueryBuilder();
                builder.InternalDate.Before(DateTime.Now);
                builder.InternalDate.Since(DateTime.Now.AddDays(-7));
                Search(client, "Arrived in the last 7 days", builder.GetQuery());

                builder = new MailQueryBuilder();
                builder.From.Contains("sender@example.com");
                Search(client, "From sender@example.com", builder.GetQuery());

                builder = new MailQueryBuilder();
                builder.From.Contains("example.com");
                Search(client, "From anyone at example.com", builder.GetQuery());

                builder = new MailQueryBuilder();
                builder.To.Contains("recipient@example.com");
                Search(client, "Sent to recipient@example.com", builder.GetQuery());

                var imapBuilder = new ImapQueryBuilder();
                imapBuilder.HasFlags(ImapMessageFlags.Keyword("custom1"));
                imapBuilder.HasNoFlags(ImapMessageFlags.Keyword("custom2"));
                Search(client, "Keyword custom1 set, custom2 not set", imapBuilder.GetQuery());
            }
        }

        private static void Search(ImapClient client, string description, MailQuery query)
        {
            var messages = client.ListMessages(query);
            Console.WriteLine($"{description,-38} {messages.Count} message(s)");
        }
    }
}
