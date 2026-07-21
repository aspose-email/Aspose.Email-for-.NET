using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.SMTP
{
    class SMTPClientActivityLogging
    {
        public static void Run()
        {
            // Build message
            MailMessage message = new MailMessage();

            // Set email address for From and TO
            message.From = "userFrom@gmail.com";
            message.To = "userTo@gmail.com";

            // Set Subject and Body
            message.Subject = "Appointment Request";
            message.Body = "Test Body";

            // Initialize SmtpClient and Set valid user name and password, Port and SecurityOptions
            SmtpClient client = new SmtpClient();
            client.Host = "smtp.gmail.com";
            client.Username = "userFrom";
            client.Password = "***********";
            client.Port = 587;
            client.SecurityOptions = SecurityOptions.SSLExplicit;
            try
            {
                client.Send(message);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
