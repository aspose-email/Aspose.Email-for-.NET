using System;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class ListingMessagesFromFolders
    {
        public static void Run()
        {
            using (var ewsClient = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var msgCollection = ewsClient.ListMessages(ewsClient.MailboxInfo.InboxUri);

                foreach (var msgInfo in msgCollection)
                {
                    Console.WriteLine(msgInfo.Subject);
                }
            }
        }
    }
}