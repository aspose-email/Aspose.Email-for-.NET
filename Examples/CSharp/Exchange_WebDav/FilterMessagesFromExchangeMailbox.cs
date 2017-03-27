using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.Dav;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Tools.Search;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.ExchangeExchange_WebDav
{
    class FilterMessagesFromExchangeMailbox
    {
        public static void Run()
        {
            // ExStart:FilterMessagesFromExchangeMailbox
            try
            {
                // Connect to Exchange Server
                const string mailboxUri = "http://exchange-server/Exchange/username";
                const string username = "username";
                const string password = "password";
                const string domain = "domain.com";

                NetworkCredential credential = new NetworkCredential(username, password, domain);
                ExchangeClient client = new ExchangeClient(mailboxUri, credential);

                // Query building by means of ExchangeQueryBuilder class
                ExchangeQueryBuilder builder = new ExchangeQueryBuilder();
                
                // Set Subject and Emails that arrived today
                builder.Subject.Contains("Newsletter");
                builder.InternalDate.On(DateTime.Now);
                
                MailQuery query = builder.GetQuery();
                // Get list of messages
                ExchangeMessageInfoCollection messages = client.ListMessages(client.MailboxInfo.InboxUri, query, false);
                Console.WriteLine("Exchange: " + messages.Count + " message(s) found.");

                // Disconnect from Exchange
                client.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            // ExEnd:FilterMessagesFromExchangeMailbox
        }        
    }
}