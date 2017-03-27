using System;
using System.Collections.Generic;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.CSharp.Exchange_EWS
{
    class PagingSupportForListingMessages
    {
        static void Run()
        {            
            ///<summary>
            /// This example shows how to list messages from Exchange Server with Paging support
            /// Introduced in Aspose.Email for .NET 6.4.0
            /// ExchangeMessagePageInfo ListMessagesByPage(string folder, int itemsPerPage);
            /// ExchangeMessagePageInfo ListMessagesByPage(string folder, int itemsPerPage, int offset);
            /// ExchangeMessagePageInfo ListMessagesByPage(string folder, int itemsPerPage, int pageOffset, ExchangeListMessagesOptions options);
            /// ExchangeMessagePageInfo ListMessagesByPage(string folder, PageInfo pageInfo);
            /// ExchangeMessagePageInfo ListMessagesByPage(string folder, PageInfo pageInfo, ExchangeListMessagesOptions options);
            ///</summary>
            // ExStart: PagingSupportForListingMessages
            using (IEWSClient client = EWSClient.GetEWSClient("exchange.domain.com", "username", "password"))
            {
                try
                {
                    // Create some test messages to be added to server for retrieval later
                    int messagesNum = 12;
                    int itemsPerPage = 5;
                    MailMessage message = null;
                    for (int i = 0; i < messagesNum; i++)
                    {
                        message = new MailMessage(
                            "from@domain.com",
                            "to@domain.com",
                            "EMAILNET-35157_1 - " + Guid.NewGuid().ToString(),
                            "EMAILNET-35157 Move paging parameters to separate class");
                        client.AppendMessage(client.MailboxInfo.InboxUri, message);
                    }
                    // Verfiy that the messages have been added to the server
                    ExchangeMessageInfoCollection totalMessageInfoCol = client.ListMessages(client.MailboxInfo.InboxUri);
                    Console.WriteLine(totalMessageInfoCol.Count);

                    ////////////////// RETREIVING THE MESSAGES USING PAGING SUPPORT ////////////////////////////////////

                    List<ExchangeMessagePageInfo> pages = new List<ExchangeMessagePageInfo>();
                    ExchangeMessagePageInfo pageInfo = client.ListMessagesByPage(client.MailboxInfo.InboxUri, itemsPerPage);
                    // Total Pages Count
                    Console.WriteLine(pageInfo.TotalCount);

                    pages.Add(pageInfo);
                    while (!pageInfo.LastPage)
                    {
                        pageInfo = client.ListMessagesByPage(client.MailboxInfo.InboxUri, itemsPerPage, pageInfo.PageOffset + 1);
                        pages.Add(pageInfo);
                    }
                    int retrievedItems = 0;
                    foreach (ExchangeMessagePageInfo pageCol in pages)
                        retrievedItems += pageCol.Items.Count;
                    // Verify total message count using paging
                    Console.WriteLine(retrievedItems);
                }
                finally
                {
                }
            }            
            // ExEnd: PagingSupportForListingMessages
        
        }
    }
}
