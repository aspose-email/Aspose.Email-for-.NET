using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    class GetMessageIdUsingImapMessageInfo
    {
        public static void Run()
        {
            // Create an imapclient with host, user and password
            ImapClient client = new ImapClient();
            client.Host = "domain.com";
            client.Username = "username";
            client.Password = "password";
            client.SelectFolder("InBox");
            ImapMessageInfoCollection msgsColl = client.ListMessages(true);
            Console.WriteLine("Total Messages: " + msgsColl.Count);
        }
    }
}