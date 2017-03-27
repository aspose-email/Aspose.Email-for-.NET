using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
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

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class CaseSensitiveEmailsFilteringUsingEWS
    {
        public static void Run()
        {
            // Connect to EWS
            const string mailboxUri = "https://outlook.office365.com/ews/exchange.asmx";
            const string username = "username";
            const string password = "password";
            const string domain = "domain";

            try
            {
                IEWSClient client = EWSClient.GetEWSClient(mailboxUri, username, password, domain);

                // ExStart:CaseSensitiveEmailsFiltering
                // Query building by means of ExchangeQueryBuilder class
                ExchangeQueryBuilder builder = new ExchangeQueryBuilder();                
                builder.Subject.Contains("Newsletter", true);
                builder.InternalDate.On(DateTime.Now);
                MailQuery query = builder.GetQuery();
                // ExEnd:CaseSensitiveEmailsFiltering

                // Get list of messages
                ExchangeMessageInfoCollection messages = client.ListMessages(client.MailboxInfo.InboxUri, query, false);
                Console.WriteLine("EWS: " + messages.Count + " message(s) found.");

                // Disconnect from EWS
                client.Dispose();                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
