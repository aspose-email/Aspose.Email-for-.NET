using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    class GetMailTips
    {
        public static void Run()
        {
            // Create instance of EWSClient class by giving credentials
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");
            Console.WriteLine("Connected to Exchange server...");
            // Provide mail tips options
            MailAddressCollection addrColl = new MailAddressCollection();
            addrColl.Add(new MailAddress("test.exchange@ex2010.local", true));
            addrColl.Add(new MailAddress("invalid.recipient@ex2010.local", true));
            GetMailTipsOptions options = new GetMailTipsOptions("administrator@ex2010.local", addrColl, MailTipsType.All);

            // Get Mail Tips
            MailTips[] tips = client.GetMailTips(options);

            // Display information about each Mail Tip
            foreach (MailTips tip in tips)
            {
                // Display Out of office message, if present
                if (tip.OutOfOffice != null)
                {
                    Console.WriteLine("Out of office: " + tip.OutOfOffice.ReplyBody.Message);
                }

                // Display the invalid email address in recipient, if present
                if (tip.InvalidRecipient == true)
                {
                    Console.WriteLine("The recipient address is invalid: " + tip.RecipientAddress);
                }
            }
        }
    }
}