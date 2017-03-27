using System;
using System.Diagnostics;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class SSLEnabledSMTPServer
    {
        public static void Run()
        {        
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "test-load.eml";

            // ExStart:SSLEnabledSMTPServer
            SmtpClient client = new SmtpClient("smtp.gmail.com");

            // Set username, Password, Port No, and SecurityOptions
            client.Username = "your.email@gmail.com";
            client.Password = "your.password";
            client.Port = 587;
            client.SecurityOptions = SecurityOptions.SSLExplicit;
            // ExEnd:SSLEnabledSMTPServer

            // Declare message as MailMessage instance
            MailMessage message = new MailMessage();

            // Use MailMessage properties like specify sender, recipient and message
            message.To = "newcustomeronnet@gmail.com";
            message.From = "newcustomeronnet@gmail.com";
            message.Subject = "Test Email";
            message.Body = "Hello World!";
            try
            {
                // Client will send this message
                client.Send(message);
                Console.WriteLine("Message sent");
            }

            catch (Exception ex)
            {
                Trace.WriteLine(ex.ToString());
            }

            Console.WriteLine(Environment.NewLine + "Email sent SSL successfully.");
        }
    }
}
