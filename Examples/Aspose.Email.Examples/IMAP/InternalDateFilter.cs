using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Tools.Search;

namespace Aspose.Email.Examples.IMAP
{
    class InternalDateFilter
    {
        public static void Run()
        {
            // Connect and log in to IMAP
            const string host = "host";
            const int port = 143;
            const string username = "user@host.com";
            const string password = "password";
            ImapClient client = new ImapClient(host, port, username, password);
            client.SelectFolder("Inbox");

            // Set conditions, Subject contains "Newsletter", Emails that arrived today
            ImapQueryBuilder builder = new ImapQueryBuilder();
            builder.Subject.Contains("Newsletter");
            builder.InternalDate.On(DateTime.Now);

            // Build the query and Get list of messages
            MailQuery query = builder.GetQuery();
            ImapMessageInfoCollection messages = client.ListMessages(query);
            foreach (ImapMessageInfo info in messages)
            {
                Console.WriteLine("Internal Date: " + info.InternalDate);
            }
            
            // Disconnect from IMAP
            client.Dispose();
        }
    }
}