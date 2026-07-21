using Aspose.Email.Clients.Imap;
using Aspose.Email.Tools.Search;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.IMAP
{
    class GetMessagesWithSpecificCriteria
    {
        public static void Run()
        {
            // Connect and log in to POP3
            const string host = "host";
            const int port = 143;
            const string username = "user@host.com";
            const string password = "password";
            ImapClient client = new ImapClient(host, port, username, password);

            try
            {

                // Emails that arrived today
                MailQueryBuilder builder = new MailQueryBuilder();
                builder.InternalDate.On(DateTime.Now);
                
                // Build the query and Get list of messages
                MailQuery query = builder.GetQuery();
                ImapMessageInfoCollection messages = client.ListMessages(query);
                Console.WriteLine("Imap: " + messages.Count + " message(s) found.");

                builder = new MailQueryBuilder();

                // Emails that arrived in last 7 days
                builder.InternalDate.Before(DateTime.Now);
                builder.InternalDate.Since(DateTime.Now.AddDays(-7));

                // Build the query and Get list of messages
                query = builder.GetQuery();
                messages = client.ListMessages(query);
                Console.WriteLine("Imap: " + messages.Count + " message(s) found.");

                builder = new MailQueryBuilder();

                // Get emails from specific sender
                builder.From.Contains("saqib.razzaq@127.0.0.1");

                // Build the query and Get list of messages
                query = builder.GetQuery();
                messages = client.ListMessages(query);
                Console.WriteLine("Imap: " + messages.Count + " message(s) found.");

                builder = new MailQueryBuilder();

                // Get emails from specific domain
                builder.From.Contains("SpecificHost.com");

                // Build the query and Get list of messages
                query = builder.GetQuery();
                messages = client.ListMessages(query);
                Console.WriteLine("Imap: " + messages.Count + " message(s) found.");

                builder = new MailQueryBuilder();

                // Get emails sent to specific recipient
                builder.To.Contains("recipient");

                ImapQueryBuilder queryBuilder = new ImapQueryBuilder();

                queryBuilder.HasFlags(ImapMessageFlags.Keyword("custom1"));

                queryBuilder.HasNoFlags(ImapMessageFlags.Keyword("custom2"));

                // Build the query and Get list of messages
                query = builder.GetQuery();
                messages = client.ListMessages(query);
                Console.WriteLine("Imap: " + messages.Count + " message(s) found.");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
