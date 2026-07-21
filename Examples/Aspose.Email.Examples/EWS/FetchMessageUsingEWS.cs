using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.EWS
{
    class FetchMessageUsingEWS
    {
        public static void Run()
        {
            try
            {
                // Create instance of ExchangeWebServiceClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

                // Call ListMessages method to list messages info from Inbox
                ExchangeMessageInfoCollection msgCollection = client.ListMessages(client.MailboxInfo.InboxUri);

                // Loop through the collection to get Message URI
                foreach (ExchangeMessageInfo msgInfo in msgCollection)
                {
	                String strMessageURI = msgInfo.UniqueUri;

	                // Now get the message details using FetchMessage()
	                MailMessage msg = client.FetchMessage(strMessageURI);
	                
                    foreach (Attachment att in msg.Attachments)
	                {
		                Console.WriteLine("Attachment Name: " + att.Name);
	                }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
