using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.Dav;
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

                // ExStart:GetEmailsWithTodayDate
                // Emails that arrived today
                MailQueryBuilder builder = new MailQueryBuilder();
                builder.InternalDate.On(DateTime.Now);
                // ExEnd:GetEmailsWithTodayDate

                // Build the query and Get list of messages
                MailQuery query = builder.GetQuery();
                ExchangeMessageInfoCollection messages = client.ListMessages(client.MailboxInfo.InboxUri, query);
                Console.WriteLine("EWS: " + messages.Count + " message(s) found.");

                builder = new MailQueryBuilder();

                // ExStart:GetEmailsOverDateRange
                // Emails that arrived in last 7 days
                builder.InternalDate.Before(DateTime.Now);
                builder.InternalDate.Since(DateTime.Now.AddDays(-7));
                // ExEnd:GetEmailsOverDateRange

                // Build the query and Get list of messages
                query = builder.GetQuery();
                messages = client.ListMessages(client.MailboxInfo.InboxUri, query);
                Console.WriteLine("EWS: " + messages.Count + " message(s) found.");

                builder = new MailQueryBuilder();

                // ExStart:GetSpecificSenderEmails
                // Get emails from specific sender
                builder.From.Contains("saqib.razzaq@127.0.0.1");
                // ExEnd:GetSpecificSenderEmails

                // Build the query and Get list of messages
                query = builder.GetQuery();
                messages = client.ListMessages(client.MailboxInfo.InboxUri, query);
                Console.WriteLine("EWS: " + messages.Count + " message(s) found.");

                builder = new MailQueryBuilder();

                // ExStart:GetSpecificDomainEmails
                // Get emails from specific domain
                builder.From.Contains("SpecificHost.com");
                // ExEnd:GetSpecificDomainEmails

                // Build the query and Get list of messages
                query = builder.GetQuery();
                messages = client.ListMessages(client.MailboxInfo.InboxUri, query);
                Console.WriteLine("EWS: " + messages.Count + " message(s) found.");

                builder = new MailQueryBuilder();

                // ExStart:GetSpecificRecipientEmails
                // Get emails sent to specific recipient
                builder.To.Contains("recipient");
                // ExEnd:GetSpecificRecipientEmails

                // Build the query and Get list of messages
                query = builder.GetQuery();
                messages = client.ListMessages(client.MailboxInfo.InboxUri, query);
                Console.WriteLine("EWS: " + messages.Count + " message(s) found.");

                // ExStart:GetSpecificMessageIdEmail
                // Get email with specific MessageId
                ExchangeQueryBuilder builder1 = new ExchangeQueryBuilder();
                builder1.MessageId.Equals("MessageID");
                // ExEnd:GetSpecificMessageIdEmail

                // Build the query and Get list of messages
                query = builder1.GetQuery();
                messages = client.ListMessages(client.MailboxInfo.InboxUri, query);
                Console.WriteLine("EWS: " + messages.Count + " message(s) found.");

                // ExStart:GetMailDeliveryNotifications
                // Get Mail Delivery Notifications
                builder1 = new ExchangeQueryBuilder();
                builder1.ContentClass.Equals(ContentClassType.MDN.ToString());
                // ExEnd:GetMailDeliveryNotifications

                //ExStart: FilterMessagesByMessageSize
                builder1 = new ExchangeQueryBuilder();
                builder1.ItemSize.Greater(80000);
                //ExEnd: FilterMessagesByMessageSize
 
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
            //ExStart: FilterMessagesWithPagingSupport
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
            //ExEnd: FilterMessagesWithPagingSupport
        }
    }
}
