// Demonstrates ODataQueryBuilder, which maps onto the OData system query options that
// Microsoft Graph understands.
//
// The point is to make the server do the work: filtering, sorting and trimming the
// returned fields cuts both the response size and the number of round trips, which is
// what keeps a listing fast on a large mailbox.

using System;
using Aspose.Email.Clients.Graph;

namespace Aspose.Email.Examples.Graph
{
    internal static class QueryGraphWithODataOptions
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
                // Sorting folders by name.
                var folderQuery = new ODataQueryBuilder { OrderBy = "displayName asc" };

                var folders = client.ListFolders(folderQuery);
                Console.WriteLine($"{folders.Count} folder(s), sorted by name:");
                foreach (var folder in folders)
                    Console.WriteLine($"  {folder.DisplayName}");

                var inbox = client.GetFolder(KnownFolders.Inbox);

                // The full set of options, as Graph names them.
                var messageQuery = new ODataQueryBuilder
                {
                    // A server-side filter expression.
                    Filter = "startswith(subject,'Re')",

                    // Sort order; "desc" reverses it.
                    OrderBy = "receivedDateTime desc",

                    // Page size, and how many to skip before the page starts.
                    Top = 10,
                    Skip = 0,

                    // Only these fields come back, which keeps the response small.
                    Select = new[] { "subject", "from", "receivedDateTime", "isRead" },

                    // Related entities to pull in with the same request.
                    Expand = new[] { "attachments" },

                    // Ask for the total count alongside the page.
                    Count = true,

                    // Full-text search. Graph does not allow Search and Filter together,
                    // so pick one or the other:
                    //   Search = "\"project update\"",

                    Format = "json"
                };

                var messages = client.ListMessages(inbox.ItemId, messageQuery);
                Console.WriteLine($"\n{messages.Count} message(s) matched:");

                foreach (MessageInfo messageInfo in messages)
                    Console.WriteLine($"  {messageInfo.Date:yyyy-MM-dd}  {messageInfo.Subject}");
            }
        }
    }
}
