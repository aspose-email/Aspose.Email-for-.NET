using System;
using System.Collections.Generic;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.CSharp.Exchange
{
    class PagingSupportForListingFolders
    {
        static void Run()
        {
            // ExStart: PagingSupportForListingFolders
            ///<summary>
            /// This example shows how to retrieve folders information from Exchange Server with Paging Support
            /// Introduced in Aspose.Email for .NET 6.4.0
            ///</summary>
            using (IEWSClient client = EWSClient.GetEWSClient("exchange.domain.com", "username", "password"))
            {
                int itemsPerPage = 5;
                ExchangeFolderInfoCollection totalFoldersCollection = client.ListSubFolders(client.MailboxInfo.RootUri);
                Console.WriteLine(totalFoldersCollection.Count);

                //////////////////// RETREIVING INFORMATION USING PAGING SUPPORT //////////////////////////////////

                List<ExchangeFolderPageInfo> pages = new List<ExchangeFolderPageInfo>();
                ExchangeFolderPageInfo pagedFoldersCollection = client.ListSubFoldersByPage(client.MailboxInfo.RootUri, itemsPerPage);

                Console.WriteLine(pagedFoldersCollection.TotalCount);

                pages.Add(pagedFoldersCollection);
                while (!pagedFoldersCollection.LastPage)
                {
                    pagedFoldersCollection = client.ListSubFoldersByPage(client.MailboxInfo.RootUri, itemsPerPage, pagedFoldersCollection.PageOffset + 1);
                    pages.Add(pagedFoldersCollection);
                }
                int retrievedFolders = 0;
                foreach (ExchangeFolderPageInfo pageCol in pages)
                    retrievedFolders += pageCol.Items.Count;

                // Verify the total count of folders
                Console.WriteLine(retrievedFolders);
            }            
            // ExEnd: PagingSupportForListingFolders
        }
    }
}
