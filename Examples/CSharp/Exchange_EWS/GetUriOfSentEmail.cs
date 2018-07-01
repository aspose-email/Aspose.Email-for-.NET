using Aspose.Email;
using Aspose.Email.Clients.Exchange.WebService;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSharp.Exchange_EWS
{
    public class GetUriOfSentEmail
    {
        public static void Run()
        {
            //ExStart: GetUriOfSentEmail
            using (IEWSClient client = EWSClient.GetEWSClient("https://exchange.office365.com/ews/exchange.asmx", "username", "password"))
            {
                // Register the event handler
                client.ItemSent += new EventHandler<SentItemEventArgs>(ItemSentHandler);

                MailMessage eml = new MailMessage("test@test.com", "test@test.com", "test message", "This is a test message");

                client.Send(eml);
            }
            //ExEnd: GetUriOfSentEmail
        }

        //ExStart: GetUriOfSentEmailSentHandler
        // Define an event handler method
        private static void ItemSentHandler(object sender, SentItemEventArgs e)
        {
            
            // Now we can get an id of sent email, which was saved in Sent Items folder
            string id = e.SentFolderItemId;
        }
        //ExEnd: GetUriOfSentEmailSentHandler
    }
}
