using System;
using System.Diagnostics;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.SMTP
{
    class ImportEML
    {
        public static void Run()
        {
            // Create an instance of the MailMessage class
            MailMessage msg = new MailMessage();

            // Import from EML format
            msg = MailMessage.Load(Data.Smtp + "test.eml", new EmlLoadOptions());

            // Create an instance of SmtpClient class
            SmtpClient client = GetSmtpClient();

            try
            {
                // Client.Send will send this message
                client.Send(msg);
                // Show Message if email sent successfully
                Console.WriteLine("Message sent");
            }

            catch (Exception ex)
            {
                Trace.WriteLine(ex.ToString());
            }

            Console.WriteLine(Environment.NewLine + "Email sent. ");
        }

        private static SmtpClient GetSmtpClient()
        {
            SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password");
            client.SecurityOptions = SecurityOptions.Auto;

            return client;
        }
    }
}
