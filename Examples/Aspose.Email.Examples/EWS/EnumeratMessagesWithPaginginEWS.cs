using System;
using System.Collections.Generic;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.EWS
{
    class EnumeratMessagesWithPaginginEWS
    {
        public static void Run()
        {
            try
            {
                // Create instance of ExchangeWebServiceClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "UserName", "Password");

                // Call ListMessages method to list messages info from Inbox
                ExchangeMessageInfoCollection msgCollection = client.ListMessages(client.GetMailboxInfo().InboxUri);
                int itemsPerPage = 5;
                List<PageInfo> pages = new List<PageInfo>();
                PageInfo pagedMessageInfoCol = client.ListMessagesByPage(client.MailboxInfo.InboxUri, itemsPerPage);
                pages.Add(pagedMessageInfoCol);
                while (!pagedMessageInfoCol.LastPage)
                {
                    pagedMessageInfoCol = client.ListMessagesByPage(client.MailboxInfo.InboxUri, itemsPerPage);
                    pages.Add(pagedMessageInfoCol);
                }
                pagedMessageInfoCol = client.ListMessagesByPage(client.MailboxInfo.InboxUri, itemsPerPage);
                while (!pagedMessageInfoCol.LastPage)
                {
                    client.ListMessages(client.MailboxInfo.InboxUri);
                }
            }
            catch (Exception ex)
            {

                Console.Write(ex.Message);
            }
            finally
            {

            }
        }
    }
}