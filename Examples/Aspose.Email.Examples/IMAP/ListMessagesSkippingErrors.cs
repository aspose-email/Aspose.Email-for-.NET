// Demonstrates how to page through a folder without one malformed message aborting the
// whole listing.
//
// With PageSettings.IgnoreExceptions set, a message the client fails to process is left
// out of the page and the failure is recorded in Items.Exceptions. Each entry is an
// ElementProcessingException whose ElementIndex tells you which position failed.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ListMessagesSkippingErrors
    {
        public static void Run()
        {
            const int itemsPerPage = 20;
            const int maxPages = 5;

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                var settings = new PageSettings
                {
                    FolderName = ImapFolderInfo.InBox,
                    IgnoreExceptions = true
                };

                var page = client.ListMessagesByPage(itemsPerPage, 0, settings);
                var pageNumber = 1;
                var failures = 0;

                while (true)
                {
                    Console.WriteLine($"Page {pageNumber}: {page.Items.Count} message(s), " +
                                      $"{page.Items.Exceptions.Count} failure(s)");

                    foreach (var error in page.Items.Exceptions)
                    {
                        failures++;
                        Console.WriteLine($"  position {error.ElementIndex}: {error.InnerException?.Message ?? error.Message}");
                    }

                    if (page.LastPage || pageNumber == maxPages)
                        break;

                    page = client.ListMessagesByPage(page.NextPage, settings);
                    pageNumber++;
                }

                Console.WriteLine($"\n{pageNumber} page(s) read, {failures} message(s) skipped.");
            }
        }
    }
}
