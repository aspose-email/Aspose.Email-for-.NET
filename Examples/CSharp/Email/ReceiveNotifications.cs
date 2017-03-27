using System;
using System.IO;
using Aspose.Email.AntiSpam;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;


/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class ReceiveNotifications
    {
        public static void Run()
        {
            try
            {
                // ExStart:ReceiveNotifications          
                // Create the message
                MailMessage msg = new MailMessage();
                msg.From = "sender@sender.com";
                msg.To = "receiver@receiver.com";
                msg.Subject = "the subject of the message";

                // Set delivery notifications for success and failed messages
                msg.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess | DeliveryNotificationOptions.OnFailure;

                // Add the MIME headers
                msg.Headers.Add_("Disposition-Notification-To", "sender@sender.com");
                msg.Headers.Add_("Disposition-Notification-To", "sender@sender.com");

                // Send the message
                SmtpClient client = new SmtpClient("host", "username", "password");
                client.Send(msg);
                // ExEnd:ReceiveNotifications
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
                
            }

        }
    }
}
