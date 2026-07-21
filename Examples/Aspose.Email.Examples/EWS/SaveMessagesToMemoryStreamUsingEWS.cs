using System;
using System.IO;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class SaveMessagesToMemoryStreamUsingEWS
    {
        public static void Run()
        {

            try
            {
                // Create instance of EWSClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

                // Call ListMessages method to list messages info from Inbox
                ExchangeMessageInfoCollection msgCollection = client.ListMessages(client.MailboxInfo.InboxUri);

                // Loop through the collection to get Message URI
                foreach (ExchangeMessageInfo msgInfo in msgCollection)
                {
                    string strMessageURI = msgInfo.UniqueUri;

                    // Now save the message in memory stream
                    MemoryStream stream = new MemoryStream();
                    client.SaveMessage(strMessageURI, stream);
                }
            }
            catch (Exception ex)
            {

                Console.Write(ex.Message);
            }
        }
    }
}