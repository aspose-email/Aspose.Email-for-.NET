using Aspose.Email.Clients.Imap;
using Aspose.Email.Tools.Search;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.IMAP
{
    class BuildingComplexQueries
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
                MailQueryBuilder builder = new MailQueryBuilder();

                // Emails from specific host, get all emails that arrived before today and all emails that arrived since 7 days ago
                builder.From.Contains("SpecificHost.com");
                builder.InternalDate.Before(DateTime.Now);
                builder.InternalDate.Since(DateTime.Now.AddDays(-7));

                // Build the query and Get list of messages
                MailQuery query = builder.GetQuery();
                ImapMessageInfoCollection messages = client.ListMessages(query);
                Console.WriteLine("Imap: " + messages.Count + " message(s) found.");

                builder = new MailQueryBuilder();

                // Specify OR condition
                builder.Or(builder.Subject.Contains("test"), builder.From.Contains("noreply@host.com"));

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
