using System;
using Aspose.Email.Clients.Pop3;
using Aspose.Email.Tools.Search;

namespace Aspose.Email.Examples.POP3
{
    class FilterMessagesFromPOP3Mailbox
    {
        public static void Run()
        {
            // Connect and log in to POP3
            const string host = "host";
            const int port = 110;
            const string username = "user@host.com";
            const string password = "password";
            Pop3Client client = new Pop3Client(host, port, username, password);

            // Set conditions, Subject contains "Newsletter", Emails that arrived today
            MailQueryBuilder builder = new MailQueryBuilder();      
            builder.Subject.Contains("Newsletter");
            builder.InternalDate.On(DateTime.Now);
            // Build the query and Get list of messages
            MailQuery query = builder.GetQuery();
            Pop3MessageInfoCollection messages = client.ListMessages(query);
            Console.WriteLine("Pop3: " + messages.Count + " message(s) found.");
        }
    }
}