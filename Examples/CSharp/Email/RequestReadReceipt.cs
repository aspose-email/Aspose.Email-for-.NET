using Aspose.Email.Clients.Smtp;
using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class RequestReadReceipt
    {
        public static void Run()
        {
            // ExStart:RequestReadReceipt
            // Create an Instance of MailMessage class
            MailMessage message = new MailMessage();

            // Specify From, To, HtmlBody, DeliveryNotificationOptions field
            message.From = "sender@sender.com";
            message.To.Add("receiver@receiver.com");
            message.HtmlBody = "<html><body>This is the Html body</body></html>";
            message.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess;
            message.Headers.Add("Return-Receipt-To", "sender@sender.com");
            message.Headers.Add("Disposition-Notification-To", "sender@sender.com");

            // Create an instance of SmtpClient Class
            SmtpClient client = new SmtpClient();

            // Specify your mailing host server, Username, Password and Port No
            client.Host = "smtp.server.com";
            client.Username = "Username";
            client.Password = "Password";
            client.Port = 25;

            try
            {
                // Client.Send will send this message
                client.Send(message);
                // Display ‘Message Sent’, only if message sent successfully
                Console.WriteLine("Message sent");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
            }
            // ExEnd:RequestReadReceipt
        }
    }
}
