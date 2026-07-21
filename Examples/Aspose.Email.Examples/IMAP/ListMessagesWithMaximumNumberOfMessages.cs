using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Tools.Search;

namespace Aspose.Email.Examples.IMAP
{
    class ListMessagesWithMaximumNumberOfMessages
    {
        public static void Run()
        {
            // Create an imapclient with host, user and password
            ImapClient client = new ImapClient("localhost", "user", "password");

            // Select the inbox folder and Get the message info collection
            ImapQueryBuilder builder = new ImapQueryBuilder();
            MailQuery query =
                builder.Or(
                builder.Or(
                builder.Or(
                builder.Or(
                builder.Subject.Contains(" (1) "),
                builder.Subject.Contains(" (2) ")),
                builder.Subject.Contains(" (3) ")),
                builder.Subject.Contains(" (4) ")),
                builder.Subject.Contains(" (5) "));
            ImapMessageInfoCollection messageInfoCol4 = client.ListMessages(query, 4);
            Console.WriteLine((messageInfoCol4.Count == 4) ? "Success" : "Failure");
        }
    }
}