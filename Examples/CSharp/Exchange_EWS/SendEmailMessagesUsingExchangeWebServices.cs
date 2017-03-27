using System;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class SendEmailMessagesUsingExchangeWebServices
    {
        public static void Run()
        {
            try
            {
                // ExStart:SendEmailMessagesUsingExchangeWebServices
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
                // ExEnd:SendEmailMessagesUsingExchangeWebServices
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
    }
}