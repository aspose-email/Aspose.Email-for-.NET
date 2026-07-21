using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Tools.Search;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.EWS
{
    class FilterMessagesUsingEWS
    {
        public static void Run()
        {
            try
            {                
                // Connect to EWS
                const string mailboxUri = "https://outlook.office365.com/ews/exchange.asmx";
                const string username = "username";
                const string password = "password";
                const string domain = "domain";

                IEWSClient client = EWSClient.GetEWSClient(mailboxUri, username, password, domain);

                // Query building by means of ExchangeQueryBuilder class
                ExchangeQueryBuilder builder = new ExchangeQueryBuilder();

                // Set Subject and Emails that arrived today
                builder.Subject.Contains("Newsletter");
                builder.InternalDate.On(DateTime.Now);

                MailQuery query = builder.GetQuery();

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
