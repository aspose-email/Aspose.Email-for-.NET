using System;
using System.Collections.Generic;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class ListingMessagesWithPagingSupport
    {
        static void Run()
        {
            // ExStart: ListingMessagesWithPagingSupport
            ///<summary>
            /// This example shows the paging support of ImapClient for listing messages from the server
            /// Available in Aspose.Email for .NET 6.4.0 and onwards
            ///</summary>
            using (ImapClient client = new ImapClient("host.domain.com", 993, "username", "password"))
            {
                try
                {
                    int messagesNum = 12;
                    int itemsPerPage = 5;
                    MailMessage message = null;
                    // Create some test messages and append these to server's inbox
                    for (int i = 0; i < messagesNum; i++)
                    {
                        message = new MailMessage(
                            "from@domain.com",
                            "to@domain.com",
                            "EMAILNET-35157 - " + Guid.NewGuid(),
                            "EMAILNET-35157 Move paging parameters to separate class");
                        client.AppendMessage(ImapFolderInfo.InBox, message);
                    }

                    // List messages from inbox
                    client.SelectFolder(ImapFolderInfo.InBox);
                    ImapMessageInfoCollection totalMessageInfoCol = client.ListMessages();
                    // Verify the number of messages added
                    Console.WriteLine(totalMessageInfoCol.Count);

                    ////////////////// RETREIVE THE MESSAGES USING PAGING SUPPORT////////////////////////////////////

                    List<ImapPageInfo> pages = new List<ImapPageInfo>();
                    ImapPageInfo pageInfo = client.ListMessagesByPage(itemsPerPage);
                    Console.WriteLine(pageInfo.TotalCount);
                    pages.Add(pageInfo);
                    while (!pageInfo.LastPage)
                    {
                        pageInfo = client.ListMessagesByPage(pageInfo.NextPage);
                        pages.Add(pageInfo);
                    }
                    int retrievedItems = 0;
                    foreach (ImapPageInfo folderCol in pages)
                        retrievedItems += folderCol.Items.Count;
                    Console.WriteLine(retrievedItems);
                }
                finally
                {
                }
            }
            // ExEnd: ListingMessagesWithPagingSupport
        }
    }
}
