using System;
using System.Configuration;
using System.Diagnostics;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class LoadSmtpConfigFile
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "EmbeddedImage.msg";

            // Declare msg as MailMessage instance
            MailMessage msg = new MailMessage();

            // Use MailMessage properties like specify sender, recipient and message
            msg.To = "asposetest123@gmail.com";
            msg.From = "aspose2@gmail.com";
            msg.Subject = "Test Email";
            msg.Body = "Hello World!";

            // ExStart:LoadSmtpConfigFile
            // Create an instance of SmtpClient class and load SMTP Authentication settings from Config file
            SmtpClient client = new SmtpClient(ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None));
            // ExEnd:LoadSmtpConfigFile

            client.SecurityOptions = SecurityOptions.Auto;

            try
            {
                // Client.Send will send this message
                client.Send(msg);
                Console.WriteLine("Message sent");
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.ToString());
            }
            Console.WriteLine(Environment.NewLine + "Message sent after loading configuration from config file.");
        }
    }
}
