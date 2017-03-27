using System;
using System.Diagnostics;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class MultipleRecipients
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "outputAttachments.msg";

            // Declare msg as MailMessage instance
            MailMessage message = new MailMessage();

            // Use MailMessage properties like specify sender, recipient and message

            // Specify the recipients mail addresses
            message.To.Add("receiver1@receiver.com");
            message.To.Add("receiver2@receiver.com");
            message.To.Add("receiver3@receiver.com");
            message.To.Add("receiver4@receiver.com");

            message.CC.Add("CC1@receiver.com");
            message.CC.Add("CC2@receiver.com");

            message.Bcc.Add("Bcc1@receiver.com");
            message.Bcc.Add("Bcc2@receiver.com");

            message.From = "newcustomeronnet@gmail.com";
            message.Subject = "Test Email";
            message.Body = "Hello World!";


            // Create an instance of SmtpClient class
            SmtpClient client = GetSmtpClient();

            try
            {
                // Client will send this message
                client.Send(message);
                // Show only if message sent successfully
                Console.WriteLine("Message sent");
            }

            catch (Exception ex)
            {
                Trace.WriteLine(ex.ToString());
            }

            Console.WriteLine(Environment.NewLine + "Email sent to multiple recipients successfully.");
        }

        private static SmtpClient GetSmtpClient()
        {
            SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password");
            client.SecurityOptions = SecurityOptions.Auto;

            return client;
        }
    }
}
