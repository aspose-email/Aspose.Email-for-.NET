// Demonstrates paging through a Graph folder and filtering with a mail query.
//
// A mailbox can hold far more messages than one response returns. ListMessages with a
// PageInfo hands back one page plus the cursor for the next, so the loop stops when
// LastPage is reached rather than by counting.

using System;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Graph;

namespace Aspose.Email.Examples.Graph
{
    internal static class ListMessagesWithPagingInGraph
    {
        public static void Run()
        {
            if (!ClientBuilder.IsGraphConfigured)
            {
                GraphExampleInfo.PrintNotConfigured();
                return;
            }

            using (var client = ClientBuilder.Graph(AuthType.ModernWithAppPermission))
            {
                // GraphQueryBuilder adds the Graph-specific fields, such as IsRead, on top
                // of the fields every mail query builder offers.
                var builder = new GraphQueryBuilder();
                builder.IsRead.Equals(false);
                var query = builder.GetQuery();

                const int itemsPerPage = 10;

                var pageInfo = client.ListMessages(KnownFolders.Inbox, new PageInfo(itemsPerPage), query);
                var page = 1;
                var total = 0;

                while (true)
                {
                    Console.WriteLine($"--- page {page} ({pageInfo.Items.Count} message(s)) ---");

                    foreach (MessageInfo messageInfo in pageInfo.Items)
                    {
                        Console.WriteLine($"  {messageInfo.Subject}");

                        // Mark it read now that it has been processed.
                        client.SetRead(messageInfo.ItemId);
                        total++;
                    }

                    if (pageInfo.LastPage)
                        break;

                    pageInfo = client.ListMessages(KnownFolders.Inbox, pageInfo.NextPage, query);
                    page++;
                }

                Console.WriteLine($"\n{total} unread message(s) processed over {page} page(s).");
            }
        }
    }
}
