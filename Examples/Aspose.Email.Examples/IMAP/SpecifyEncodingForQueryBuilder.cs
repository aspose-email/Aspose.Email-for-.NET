using Aspose.Email.Clients.Imap;
using Aspose.Email.Tools.Search;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.IMAP
{
    class SpecifyEncodingForQueryBuilder
    {
        public static void Run()
        {
            // Connect and log in to IMAP
            const string host = "host";
            const int port = 143;
            const string username = "user@host.com";
            const string password = "password";
            ImapClient client = new ImapClient(host, port, username, password);

            try
            {
                client.SelectFolder("Inbox");

                // Set conditions
                ImapQueryBuilder builder = new ImapQueryBuilder(Encoding.UTF8);
                builder.Subject.Contains("ğüşıöç", true);
                MailQuery query = builder.GetQuery();

                // Get list of messages
                ImapMessageInfoCollection messages = client.ListMessages(query);
                foreach (ImapMessageInfo info in messages)
                {
                    Console.WriteLine("Message Id: " + info.MessageId);
                }

                // Disconnect from IMAP
                client.Dispose();           
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
