using Aspose.Email.Clients.Smtp;
using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.SMTP
{
    class SendingEmailWithAlternateText
    {
        public static void Run()
        {
            // Create an instance of MailMessage class
            MailMessage message = new MailMessage();

            // From and To field
            message.From = "sender@sender.com";
            message.To.Add("receiver@receiver.com");

            AlternateView alternate;

            // Create an instance of AlternateView to view an email message using the content specified in the string
            alternate = AlternateView.CreateAlternateViewFromString("This is the alternate Text");

            // Add alternate text
            message.AlternateViews.Add(alternate);

            // Create an instance of SmtpClient Class
            SmtpClient client = new SmtpClient();

            // Specify your mailing host server, user name, mail password and Port #
            client.Host = "smtp.server.com";
            client.Username = "Username";
            client.Password = "Password";
            client.Port = 25;
            try
            {
                // Client.Send will send this message
                client.Send(message);
            }

            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
            }

        }
    }
}
