using System;
using System.Diagnostics;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class ImportEML
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "test.eml";

            // Create an instance of the MailMessage class
            MailMessage msg = new MailMessage();

            // Import from EML format
            msg = MailMessage.Load(dstEmail, new EmlLoadOptions());

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

            Console.WriteLine(Environment.NewLine + "Email sent. " + dstEmail);
        }

        private static SmtpClient GetSmtpClient()
        {
            SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password");
            client.SecurityOptions = SecurityOptions.Auto;

            return client;
        }
    }
}
