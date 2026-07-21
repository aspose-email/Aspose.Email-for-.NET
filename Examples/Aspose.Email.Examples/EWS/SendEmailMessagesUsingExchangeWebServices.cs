using System;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.EWS
{
    class SendEmailMessagesUsingExchangeWebServices
    {
        public static void Run()
        {
            try
            {
                // Create instance of IEWSClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

                // Create instance of type MailMessage
                MailMessage msg = new MailMessage();
                msg.From = "sender@domain.com";
                msg.To = "recipient@ domain.com ";
                msg.Subject = "Sending message from exchange server";
                msg.HtmlBody = "<h3>sending message from exchange server</h3>";

                // Send the message
                client.Send(msg);
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
    }
}