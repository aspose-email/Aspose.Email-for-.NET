// Demonstrates capping the number of messages a search returns.
//
// ListMessages(query, maxNumberOfMessages) stops after the given number of matches, so
// a broad search on a large folder does not download thousands of summaries when you
// only need a few - for example to check whether anything matches at all.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ListMessagesWithMaximumNumberOfMessages
    {
        public static void Run()
        {
            const int maxMessages = 5;

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                var builder = new ImapQueryBuilder();
                builder.InternalDate.Since(DateTime.Today.AddDays(-30));
                var query = builder.GetQuery();

                var all = client.ListMessages(query);
                var capped = client.ListMessages(query, maxMessages);

                Console.WriteLine($"Arrived in the last 30 days: {all.Count} message(s)");
                Console.WriteLine($"With a cap of {maxMessages}:           {capped.Count} message(s)");

                foreach (var info in capped)
                    Console.WriteLine($"  {info.InternalDate:g}  {info.Subject}");
            }
        }
    }
}
