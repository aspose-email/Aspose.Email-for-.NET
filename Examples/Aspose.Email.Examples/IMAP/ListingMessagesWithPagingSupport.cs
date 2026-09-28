// Demonstrates how to list a folder page by page instead of all at once.
//
// ListMessagesByPage returns one page of message summaries together with paging
// information: TotalCount, LastPage and NextPage. Pass the offset of NextPage back in to
// get the following page, until LastPage is true. PageSettings chooses the folder and
// the sort direction.
//
// The example appends 12 messages to a uniquely named folder, reads them five at a time,
// and deletes the folder at the end.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ListingMessagesWithPagingSupport
    {
        public static void Run()
        {
            const int messageCount = 12;
            const int itemsPerPage = 5;

            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.CreateFolder(folderName);

                try
                {
                    for (var i = 1; i <= messageCount; i++)
                    {
                        client.AppendMessage(folderName,
                            new MailMessage("from@example.com", "to@example.com", $"Message {i:D2}", "Body"));
                    }

                    client.SelectFolder(folderName);

                    var settings = new PageSettings { FolderName = folderName };
                    var page = client.ListMessagesByPage(itemsPerPage, 0, settings);
                    Console.WriteLine($"{page.TotalCount} message(s) in '{folderName}', {itemsPerPage} per page.");

                    var pageNumber = 1;
                    var retrieved = 0;

                    while (true)
                    {
                        Console.WriteLine($"\nPage {pageNumber} (offset {page.PageOffset}):");
                        foreach (var info in page.Items)
                            Console.WriteLine("  " + info.Subject);

                        retrieved += page.Items.Count;

                        if (page.LastPage)
                            break;

                        page = client.ListMessagesByPage(itemsPerPage, page.NextPage.PageOffset, settings);
                        pageNumber++;
                    }

                    Console.WriteLine($"\nRetrieved {retrieved} message(s) in {pageNumber} page(s).");
                }
                finally
                {
                    client.DeleteFolder(folderName);
                }
            }
        }
    }
}
