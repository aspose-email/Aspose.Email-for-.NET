using System;
using System.Collections.Generic;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;
using Aspose.Email.Tools.Search;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class SearchWithPagingSupport
    {
        static void Run()
        { 
            // ExStart:SearchWithPagingSupport
            ///<summary>
            /// This example shows how to search for messages using ImapClient of the API with paging support
            /// Introduced in Aspose.Email for .NET 6.4.0
            ///</summary>
            using (ImapClient client = new ImapClient("host.domain.com", 84, "username", "password"))
            {
                try
                {
                    // Append some test messages
                    int messagesNum = 12;
                    int itemsPerPage = 5;
                    MailMessage message = null;
                    for (int i = 0; i < messagesNum; i++)
                    {
                        message = new MailMessage(
                            "from@domain.com",
                            "to@domain.com",
                            "EMAILNET-35128 - " + Guid.NewGuid(),
                            "111111111111111");
                        client.AppendMessage(ImapFolderInfo.InBox, message);
                    }
                    string body = "2222222222222";
                    for (int i = 0; i < messagesNum; i++)
                    {
                        message = new MailMessage(
                            "from@domain.com",
                            "to@domain.com",
                            "EMAILNET-35128 - " + Guid.NewGuid(),
                            body);
                        client.AppendMessage(ImapFolderInfo.InBox, message);
                    }

                    ImapQueryBuilder iqb = new ImapQueryBuilder();
                    iqb.Body.Contains(body);
                    MailQuery query = iqb.GetQuery();

                    client.SelectFolder(ImapFolderInfo.InBox);
                    ImapMessageInfoCollection totalMessageInfoCol = client.ListMessages(query);
                    Console.WriteLine(totalMessageInfoCol.Count);

                    //////////////////////////////////////////////////////

                    List<ImapPageInfo> pages = new List<ImapPageInfo>();
                    ImapPageInfo pageInfo = client.ListMessagesByPage(ImapFolderInfo.InBox, query, itemsPerPage);
                    pages.Add(pageInfo);
                    while (!pageInfo.LastPage)
                    {
                        pageInfo = client.ListMessagesByPage(ImapFolderInfo.InBox, query, pageInfo.NextPage);
                        pages.Add(pageInfo);
                    }
                    int retrievedItems = 0;
                    foreach (ImapPageInfo folderCol in pages)
                        retrievedItems += folderCol.Items.Count;
                }
                finally
                {
                }
            }
            // ExEnd: SearchWithPagingSupport
        }
    }
}
