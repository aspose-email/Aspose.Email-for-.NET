using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.Dav;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Tools.Search;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.EWS
{
    class FilterMessagesOnCriteriaUsingEWS
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

                // Emails that arrived today
                MailQueryBuilder builder = new MailQueryBuilder();
                builder.InternalDate.On(DateTime.Now);

                // Build the query and Get list of messages
                MailQuery query = builder.GetQuery();
                ExchangeMessageInfoCollection messages = client.ListMessages(client.MailboxInfo.InboxUri, query);
                Console.WriteLine("EWS: " + messages.Count + " message(s) found.");

                builder = new MailQueryBuilder();

                // Emails that arrived in last 7 days
                builder.InternalDate.Before(DateTime.Now);
                builder.InternalDate.Since(DateTime.Now.AddDays(-7));

                // Build the query and Get list of messages
                query = builder.GetQuery();
                messages = client.ListMessages(client.MailboxInfo.InboxUri, query);
                Console.WriteLine("EWS: " + messages.Count + " message(s) found.");

                builder = new MailQueryBuilder();

                // Get emails from specific sender
                builder.From.Contains("saqib.razzaq@127.0.0.1");

                // Build the query and Get list of messages
                query = builder.GetQuery();
                messages = client.ListMessages(client.MailboxInfo.InboxUri, query);
                Console.WriteLine("EWS: " + messages.Count + " message(s) found.");

                builder = new MailQueryBuilder();

                // Get emails from specific domain
                builder.From.Contains("SpecificHost.com");

                // Build the query and Get list of messages
                query = builder.GetQuery();
                messages = client.ListMessages(client.MailboxInfo.InboxUri, query);
                Console.WriteLine("EWS: " + messages.Count + " message(s) found.");

                builder = new MailQueryBuilder();

                // Get emails sent to specific recipient
                builder.To.Contains("recipient");

                // Build the query and Get list of messages
                query = builder.GetQuery();
                messages = client.ListMessages(client.MailboxInfo.InboxUri, query);
                Console.WriteLine("EWS: " + messages.Count + " message(s) found.");

                // Get email with specific MessageId
                ExchangeQueryBuilder builder1 = new ExchangeQueryBuilder();
                builder1.MessageId.Equals("MessageID");

                // Build the query and Get list of messages
                query = builder1.GetQuery();
                messages = client.ListMessages(client.MailboxInfo.InboxUri, query);
                Console.WriteLine("EWS: " + messages.Count + " message(s) found.");

                // Get Mail Delivery Notifications
                builder1 = new ExchangeQueryBuilder();
                builder1.ContentClass.Equals(ContentClassType.MDN.ToString());

                builder1 = new ExchangeQueryBuilder();
                builder1.ItemSize.Greater(80000);
 
                // Build the query and Get list of messages
                query = builder1.GetQuery();
                messages = client.ListMessages(client.MailboxInfo.InboxUri, query);
                Console.WriteLine("EWS: " + messages.Count + " message(s) found.");

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static void FilterMessagesWithPagingSupport()
        {
            const string mailboxUri = "https://outlook.office365.com/ews/exchange.asmx";
            const string username = "username";
            const string password = "password";
            const string domain = "domain";

            IEWSClient client = EWSClient.GetEWSClient(mailboxUri, username, password, domain);
            int itemsPerPage = 5;
            string sGuid = Guid.NewGuid().ToString();
            string str1 = sGuid + " - " + "Query 1";
            string str2 = sGuid + " - " + "Query 2";

            MailQueryBuilder queryBuilder1 = new MailQueryBuilder();
            queryBuilder1.Subject.Contains(str1);
            MailQuery query1 = queryBuilder1.GetQuery();
            List<ExchangeMessagePageInfo> pages = new List<ExchangeMessagePageInfo>();
            ExchangeMessagePageInfo pageInfo = client.ListMessagesByPage(client.MailboxInfo.InboxUri, query1, itemsPerPage);
            pages.Add(pageInfo);
            int str1Count = pageInfo.Items.Count;
            while (!pageInfo.LastPage)
            {
                pageInfo = client.ListMessagesByPage(client.MailboxInfo.InboxUri, query1, itemsPerPage, pageInfo.PageOffset + 1);
                pages.Add(pageInfo);
                str1Count += pageInfo.Items.Count;
            }
        }
    }
}
