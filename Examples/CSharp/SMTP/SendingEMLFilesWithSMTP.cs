using System;
using System.Diagnostics;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.SMTP
{
    class SendingEMLFilesWithSMTP
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "Message.eml";

            // Create an instance of the MailMessage class
            MailMessage message = new MailMessage();

            // Import from EML format
            message = MailMessage.Load(dstEmail, new EmlLoadOptions());

            // Create an instance of SmtpClient class
            SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password");
            client.SecurityOptions = SecurityOptions.Auto;

            try
            {
                // Client.Send will send this message
                client.Send(message);
                Console.WriteLine("Message sent");
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.ToString());
            }
            Console.WriteLine(Environment.NewLine + "Email sent using EML file successfully. " + dstEmail);
        }
    }
}
