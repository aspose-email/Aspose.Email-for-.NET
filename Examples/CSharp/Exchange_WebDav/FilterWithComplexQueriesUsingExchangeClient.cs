using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.Dav;
using Aspose.Email.Tools.Search;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_WebDav
{
    class FilterWithComplexQueriesUsingExchangeClient
    {
        public static void Run()
        {
            try
            {
                // Create instance of ExchangeClient class by giving credentials
                ExchangeClient client = new ExchangeClient("http://ex07sp1/exchange/Administrator", "user", "pwd", "domain");
                MailQueryBuilder builder = new MailQueryBuilder();

                // ExStart:CombineQueriesWithAND
                // Emails from specific host, get all emails that arrived before today and all emails that arrived since 7 days ago
                builder.From.Contains("SpecificHost.com");
                builder.InternalDate.Before(DateTime.Now);
                builder.InternalDate.Since(DateTime.Now.AddDays(-7));
                // ExEnd:CombineQueriesWithAND

                MailQuery query = builder.GetQuery();
                ExchangeMessageInfoCollection messages = client.ListMessages(client.MailboxInfo.InboxUri, query, false);
                Console.WriteLine("Exchange: " + messages.Count + " message(s) found.");

                builder = new MailQueryBuilder();

                // ExStart:CombiningQueriesWithOR
                // Specify OR condition
                builder.Or(builder.Subject.Contains("test"), builder.From.Contains("noreply@host.com"));
                // ExEnd:CombiningQueriesWithOR

                // Build the query and Get list of messages
                query = builder.GetQuery();
                messages = client.ListMessages(client.MailboxInfo.InboxUri, query, false);
                Console.WriteLine("Exchange: " + messages.Count + " message(s) found.");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
