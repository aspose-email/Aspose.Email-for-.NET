using System;
using System.Collections.Generic;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class EnumeratMessagesWithPaginginEWS
    {
        public static void Run()
        {
            try
            {
                // ExStart:EnumeratMessagesWithPaginginEWS
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
            // ExEnd:EnumeratMessagesWithPaginginEWS
        }
    }
}