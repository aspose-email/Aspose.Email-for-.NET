// Demonstrates listing a folder page by page with explicit page offsets.
//
// ListMessagesByPage(folder, itemsPerPage, pageOffset) returns the requested page with
// TotalCount and LastPage; increase the offset until LastPage is true. The example
// appends 12 messages to a uniquely named folder, reads them five at a time, and
// deletes the folder at the end.

using System;
using System.Linq;

namespace Aspose.Email.Examples.EWS
{
    internal static class PagingSupportForListingMessages
    {
        public static void Run()
        {
            const int itemsPerPage = 5;
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.InboxUri, folderName);

                try
                {
                    var messages = Enumerable.Range(1, 12)
                        .Select(i => new MailMessage("sender@example.com", "receiver@example.com",
                            $"Message {i:D2}", "Body"));
                    client.AppendMessages(folder.Uri, messages);

                    var page = client.ListMessagesByPage(folder.Uri, itemsPerPage);
                    Console.WriteLine($"{page.TotalCount} message(s), {itemsPerPage} per page.");

                    var retrieved = 0;
                    while (true)
                    {
                        var subjects = string.Join(", ", page.Items.Select(i => i.Subject));
                        Console.WriteLine($"Page {page.PageOffset}: {subjects}");
                        retrieved += page.Items.Count;

                        if (page.LastPage)
                            break;

                        page = client.ListMessagesByPage(folder.Uri, itemsPerPage, page.PageOffset + 1);
                    }

                    Console.WriteLine($"Retrieved {retrieved} message(s).");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
