using System.IO;
using System;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;
using Aspose.Email.Pop3;
using Aspose.Email;
using Aspose.Email.Mime;
using Aspose.Email.Imap;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class SettingTextBody
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "test.eml";

            // Declare msg as MailMessage instance
            MailMessage msg = new MailMessage();

            // Use MailMessage properties like specify sender, recipient and message
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
                // Show only if message sent successfully
                Console.WriteLine("Message sent");
            }

            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
            }

            Console.WriteLine(Environment.NewLine + "Email sent with Text body.");
        }

        private static SmtpClient GetSmtpClient()
        {
            SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password");
            client.SecurityOptions = SecurityOptions.Auto;

            return client;
        }
    }
}
