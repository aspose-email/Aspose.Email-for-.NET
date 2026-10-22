// Demonstrates walking through a large folder page by page with PageInfo.
//
// Pass a PageInfo with the page size to ListMessagesByPage; every result carries
// NextPage, which you pass back in until LastPage is true. Only one page of summaries
// is held in memory at a time. PagingSupportForListingMessages shows the variant with
// explicit page offsets.

using System;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.EWS
{
    internal static class EnumeratMessagesWithPaginginEWS
    {
        public static void Run()
        {
            const int itemsPerPage = 10;
            const int maxPages = 5;

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var inboxUri = client.MailboxInfo.InboxUri;

                var page = client.ListMessagesByPage(inboxUri, new PageInfo(itemsPerPage));
                Console.WriteLine($"{page.TotalCount} message(s) in the Inbox.");

                var pageNumber = 1;
                while (true)
                {
                    Console.WriteLine($"\nPage {pageNumber}:");
                    foreach (var info in page.Items)
                        Console.WriteLine($"  {info.InternalDate:g}  {info.Subject}");

                    if (page.LastPage || pageNumber == maxPages)
                        break;

                    page = client.ListMessagesByPage(inboxUri, page.NextPage);
                    pageNumber++;
                }
            }
        }
    }
}
