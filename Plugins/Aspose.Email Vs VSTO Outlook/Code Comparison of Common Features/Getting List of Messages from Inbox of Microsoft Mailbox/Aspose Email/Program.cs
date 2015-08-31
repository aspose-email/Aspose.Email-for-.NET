using Aspose.Email.Exchange;
using System;

namespace Aspose_Email
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create instance of ExchangeClient class by giving credentials
            ExchangeClient client = new ExchangeClient("http://MachineName/exchange/Username",
                            "username", "password", "domain");
            // Call ListMessages method to list messages info from Inbox
            ExchangeMessageInfoCollection msgCollection = client.ListMessages(client.MailboxInfo.InboxUri);

            // Loop through the collection to display the basic information
            foreach (ExchangeMessageInfo msgInfo in msgCollection)
            {
                Console.WriteLine("Subject: " + msgInfo.Subject);
                Console.WriteLine("From: " + msgInfo.From.ToString());
                Console.WriteLine("To: " + msgInfo.To.ToString());
                Console.WriteLine("Message ID: " + msgInfo.MessageId);
                Console.WriteLine("Unique URI: " + msgInfo.UniqueUri);
                Console.WriteLine("==================================");
            }
        }
    }
}
