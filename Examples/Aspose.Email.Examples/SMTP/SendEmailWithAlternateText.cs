using System;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;

namespace Aspose.Email.Examples.SMTP
{
    class SendEmailWithAlternateText
    {
        public static void Run()
        {
            //Create an instance of the MailMessage class
            MailMessage message = new MailMessage();

            // Set From field, To field and Plain text body
            message.From = "sender@sender.com";
            message.To.Add("receiver@receiver.com");
            message.Body = "This is Plain Text Body";

            // Create an instance of the SmtpClient class
            SmtpClient client = new SmtpClient();

            // And Specify your mailing host server, Username, Password and Port
            client.Host = "smtp.server.com";
            client.Username = "Username";
            client.Password = "Password";
            client.Port = 25;

            try
            {
                // Client.Send will send this message
                client.Send(message);
                Console.WriteLine("Message sent");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
            }

            Console.WriteLine("Press enter to quit");
            Console.Read();
        }
    }
}