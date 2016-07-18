using System.Collections.Generic;
using Aspose.Email.Exchange;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange
{
    class EnumeratMessagesWithPaginginEWS
    {
        public static void Run()
        {
            // ExStart:EnumeratMessagesWithPaginginEWS
            // Create instance of ExchangeWebServiceClient class by giving credentials
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "Shabir.haider@studentpartner.com", "LoveAir1993");

            // Call ListMessages method to list messages info from Inbox
            ExchangeMessageInfoCollection msgCollection = client.ListMessages(client.GetMailboxInfo().InboxUri);

            try
            {
                int itemsPerPage = 5;

                List<ExchangeMessageInfoCollection> pages = new List<ExchangeMessageInfoCollection>();

                ExchangeMessageInfoCollection pagedMessageInfoCol = client.ListMessages(client.MailboxInfo.InboxUri, itemsPerPage);

                pages.Add(pagedMessageInfoCol);

                while (!pagedMessageInfoCol.LastPage)
                {
                    pagedMessageInfoCol = client.ListMessages(client.MailboxInfo.InboxUri, itemsPerPage, pagedMessageInfoCol.LastItemOffset.Value + 1);
                    pages.Add(pagedMessageInfoCol);
                }

                pagedMessageInfoCol = client.ListMessages(client.MailboxInfo.InboxUri, itemsPerPage);

                while (!pagedMessageInfoCol.LastPage)
                {
                    client.ListMessages(client.MailboxInfo.InboxUri, pagedMessageInfoCol, itemsPerPage, pagedMessageInfoCol.LastItemOffset.Value + 1);
                }
            }
            finally
            {

            }
            // ExEnd:EnumeratMessagesWithPaginginEWS
        }
    }
}