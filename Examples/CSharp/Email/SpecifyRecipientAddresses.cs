using Aspose.Email.Clients.Smtp;
using Aspose.Email.Mime;
using System;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class SpecifyRecipientAddresses
    {
        public static void Run()
        {
            // ExStart:SpecifyRecipientAddresses
            // Declare message as MailMessage instance
            // Create an Instance of MailMessage class
            MailMessage message = new MailMessage();

            // Specify From address
            message.From = "sender@sender.com";

            //  Specify recipients’ mail addresses
            message.To.Add("receiver1@receiver.com");
            message.To.Add("receiver2@receiver.com");
            message.To.Add("receiver3@receiver.com");
            
            // Specify CC addresses
            message.CC.Add("CC1@receiver.com");
            message.CC.Add("CC2@receiver.com");

            // Specify BCC addresses
            message.Bcc.Add("Bcc1@receiver.com");
            message.Bcc.Add("Bcc2@receiver.com");

            // Create an instance of SmtpClient Class
            SmtpClient client = new SmtpClient();

            // Specify your mailing host server, Username, Password, Port
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
            // ExEnd:SpecifyRecipientAddresses
        }
    }
}
