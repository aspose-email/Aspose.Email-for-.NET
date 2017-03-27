using System;
using System.Diagnostics;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class SettingTextBody
    {
        // ExStart:SettingTextBody
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "test.eml";

            // Declare msg as MailMessage instance
            MailMessage msg = new MailMessage();

            // Use MailMessage properties like specify sender, recipient and message
            msg.From = "newcustomeronnet@gmail.com";
            msg.To = "newcustomeronnet2@gmail.com";
            msg.Subject = "Test subject";
            msg.Body = "This is text body";
            SmtpClient client = GetSmtpClient();
            try
            {
                // Client will send this message
                client.Send(msg);
                Console.WriteLine("Message sent");
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.ToString());
            }
            Console.WriteLine(Environment.NewLine + "Email sent with Text body.");
        }

        private static SmtpClient GetSmtpClient()
        {
            SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password");
            client.SecurityOptions = SecurityOptions.Auto;
            return client;
        }
        // ExEnd:SettingTextBody
    }
}
