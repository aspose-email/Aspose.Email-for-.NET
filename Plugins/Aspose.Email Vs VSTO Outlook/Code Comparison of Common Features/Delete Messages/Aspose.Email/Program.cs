using Aspose.Email.Exchange;
using System;
using System.Collections.Generic;
using System.Text;

namespace Aspose.Email
{
    class Program
    {
        static void Main(string[] args)
        {
            string MailBoxURI = "http://MachineName/exchange/Username";
            string UserName = "username";
            string Password = "password";
            string Domain = "domain";

            // Create instance of ExchangeClient class by giving credentials
            ExchangeClient client = new ExchangeClient(MailBoxURI, UserName, Password, Domain);
            // Call ListMessages method to list messages info from Inbox
            ExchangeMessageInfoCollection msgCollection = client.ListMessages(client.MailboxInfo.InboxUri);

            // Get URI of Message to Delete
            string MessageURI= msgCollection[0].UniqueUri;

            // Delete the message
            client.DeleteMessage(MessageURI);
        }
    }
}
