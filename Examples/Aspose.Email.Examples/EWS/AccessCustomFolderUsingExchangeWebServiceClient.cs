using System;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class AccessCustomFolderUsingExchangeWebServiceClient
    {
        public static void Run()
        {
            // Create instance of EWSClient class by giving credentials
            var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission);

            // Create ExchangeMailboxInfo, ExchangeMessageInfoCollection instance
            var mailbox = client.GetMailboxInfo();
            ExchangeMessageInfoCollection messages;
            ExchangeFolderInfo subfolderInfo;

            // Check if specified custom folder exisits and Get all the messages info from the target Uri
            client.FolderExists(mailbox.InboxUri, "TestInbox", out subfolderInfo);

            //if custom folder exists
            if (subfolderInfo != null)
            {
                messages = client.ListMessages(subfolderInfo.Uri);

                // Parse all the messages info collection
                foreach (var messageInfo in messages)
                {
                    // now get the message details using FetchMessage()
                    var msg = client.FetchMessage(messageInfo.UniqueUri);
                    Console.WriteLine($"Subject: {msg.Subject}");
                }
            }
            else
            {
                Console.WriteLine("No folder with this name found.");
            }
        }
    }
}