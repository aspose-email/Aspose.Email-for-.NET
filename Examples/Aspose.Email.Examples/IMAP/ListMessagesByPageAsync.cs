// Demonstrates paging through search results with the task-based API.
//
// The first call takes a PageInfo with the page size. Every result carries NextPage,
// which you pass back in until LastPage is true, so only one page of message summaries
// is held in memory at a time.

using System;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ListMessagesByPageAsync
    {
        public static void Run()
        {
            RunAsync().GetAwaiter().GetResult();
        }

        private static async Task RunAsync()
        {
            const int itemsPerPage = 10;
            const int maxPages = 5;

            using (var imapClient = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            using (var cancellation = new CancellationTokenSource(TimeSpan.FromMinutes(2)))
            {
                IAsyncImapClient client = imapClient;

                await client.SelectFolderAsync(ImapFolderInfo.InBox, token: cancellation.Token);

                var builder = new ImapQueryBuilder();
                builder.InternalDate.Since(DateTime.Today.AddDays(-30));
                var query = builder.GetQuery();

                var settings = new PageSettings { FolderName = ImapFolderInfo.InBox };

                var page = await client.ListMessagesByPageAsync(query, new PageInfo(itemsPerPage), settings, cancellation.Token);
                var pageNumber = 1;

                Console.WriteLine($"{page.TotalCount} message(s) arrived in the last 30 days.");

                while (true)
                {
                    Console.WriteLine($"\nPage {pageNumber}:");
                    foreach (var info in page.Items)
                        Console.WriteLine($"  {info.InternalDate:g}  {info.Subject}");

                    if (page.LastPage || pageNumber == maxPages)
                        break;

                    page = await client.ListMessagesByPageAsync(query, page.NextPage, settings, cancellation.Token);
                    pageNumber++;
                }
            }
        }
    }
}
