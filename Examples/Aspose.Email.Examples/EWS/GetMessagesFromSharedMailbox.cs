using System;
using System.Net;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.EWS
{
    public class GetMessagesFromSharedMailbox
    {
        public static void Run()
        {
            const string mailboxUri = "<HOST>";
            const string domain = "";
            const string username = "<EMAIL ADDRESS>";
            const string password = "<PASSWORD>";
            const string sharedEmail = "<SHARED EMAIL ADDRESS>";
            NetworkCredential credentials = new NetworkCredential(username, password, domain);
            IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credentials);

            string[] items = client.ListItems(sharedEmail, "Inbox");

            foreach (string item in items)
            {
                MapiMessage msg = client.FetchItem(item);
                Console.WriteLine("Subject:" + msg.Subject);
            }
            client.Dispose();
            Console.WriteLine("GetMessagesFromSharedMailbox executed successfully");
        }
    }
}