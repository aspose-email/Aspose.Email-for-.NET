// Demonstrates paging through the results of a search, using the task-based API.
//
// ListMessagesByPageAsync takes the search query, a PageInfo with the page size and the
// PageSettings. Only matching messages are paged, so TotalCount is the number of hits;
// each result carries NextPage to pass back in until LastPage is true.
//
// The example appends 12 messages of two kinds to a uniquely named folder, pages through
// the ones whose body contains a marker, and deletes the folder at the end. Servers that
// index message bodies in the background may need a moment before new messages match.

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class SearchWithPagingSupport
    {
        public static void Run()
        {
            RunAsync().GetAwaiter().GetResult();
        }

        private static async Task RunAsync()
        {
            const int messagesPerKind = 6;
            const int itemsPerPage = 4;
            const string marker = "quarterly-report";

            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var imapClient = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                IAsyncImapClient client = imapClient;

                await client.CreateFolderAsync(folderName);

                try
                {
                    var messages = new List<MailMessage>();
                    for (var i = 1; i <= messagesPerKind; i++)
                    {
                        messages.Add(new MailMessage("from@example.com", "to@example.com",
                            $"Report {i}", $"Attached is the {marker} for region {i}."));
                        messages.Add(new MailMessage("from@example.com", "to@example.com",
                            $"Chat {i}", "Lunch at noon?"));
                    }

                    await client.AppendMessagesAsync(messages, folderName);
                    await client.SelectFolderAsync(folderName);

                    var builder = new ImapQueryBuilder();
                    builder.Body.Contains(marker);
                    var query = builder.GetQuery();

                    var settings = new PageSettings { FolderName = folderName };
                    var page = await client.ListMessagesByPageAsync(query, new PageInfo(itemsPerPage), settings);

                    Console.WriteLine($"{page.TotalCount} of {messages.Count} message(s) mention '{marker}', " +
                                      $"{itemsPerPage} per page.");

                    var pageNumber = 1;
                    while (true)
                    {
                        Console.WriteLine($"\nPage {pageNumber}:");
                        foreach (var info in page.Items)
                            Console.WriteLine("  " + info.Subject);

                        if (page.LastPage)
                            break;

                        page = await client.ListMessagesByPageAsync(query, page.NextPage, settings);
                        pageNumber++;
                    }
                }
                finally
                {
                    await client.DeleteFolderAsync(folderName);
                }
            }
        }
    }
}
