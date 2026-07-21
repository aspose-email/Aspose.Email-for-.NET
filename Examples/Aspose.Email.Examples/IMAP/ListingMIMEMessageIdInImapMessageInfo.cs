using Aspose.Email.Clients.Imap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.IMAP
{
    class ListingMIMEMessageIdInImapMessageInfo
    {
        public static void Run()
        {
            ImapClient client = new ImapClient();
            client.Host = "domain.com";
            client.Username = "username";
            client.Password = "password";
                
            try
            {
                ImapMessageInfoCollection messageInfoCol = client.ListMessages("Inbox");
                foreach (ImapMessageInfo info in messageInfoCol)
                {
                    // Display MIME Message ID
                    Console.WriteLine("Message Id = " + info.MessageId);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
