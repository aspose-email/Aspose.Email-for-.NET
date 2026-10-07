// Demonstrates filtering a folder on the server instead of downloading it.
//
// Build the conditions with ImapQueryBuilder, turn them into a MailQuery with GetQuery
// and pass that to ListMessages: the server runs the IMAP SEARCH and only the matching
// summaries come back. All conditions in one builder must hold at the same time.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class FilteringMessagesFromIMAPMailbox
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                var builder = new ImapQueryBuilder();
                builder.Subject.Contains("newsletter");
                builder.InternalDate.Since(DateTime.Today.AddDays(-30));

                var messages = client.ListMessages(builder.GetQuery());

                Console.WriteLine($"'newsletter' in the subject, last 30 days: {messages.Count} message(s)");
                foreach (var info in messages)
                    Console.WriteLine($"  {info.InternalDate:d}  {info.From}  {info.Subject}");
            }
        }
    }
}
