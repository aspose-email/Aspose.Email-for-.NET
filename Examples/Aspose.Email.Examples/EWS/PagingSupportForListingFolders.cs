// Demonstrates listing the subfolders of a folder page by page.
//
// ListSubFoldersByPage returns one page of folders together with TotalCount and
// LastPage; pass the next page offset back in until LastPage is true. Paging matters
// for public folder trees or archive mailboxes with thousands of folders.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class PagingSupportForListingFolders
    {
        public static void Run()
        {
            const int itemsPerPage = 5;

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var rootUri = client.MailboxInfo.RootUri;

                var page = client.ListSubFoldersByPage(rootUri, itemsPerPage);
                Console.WriteLine($"{page.TotalCount} top-level folder(s), {itemsPerPage} per page.");

                var pageNumber = 1;
                while (true)
                {
                    Console.WriteLine($"\nPage {pageNumber}:");
                    foreach (var folder in page.Items)
                        Console.WriteLine("  " + folder.DisplayName);

                    if (page.LastPage)
                        break;

                    page = client.ListSubFoldersByPage(rootUri, itemsPerPage, page.PageOffset + 1);
                    pageNumber++;
                }
            }
        }
    }
}
