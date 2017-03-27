using System;
using System.Diagnostics;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class CustomizingEmailHeader
    {
        public static void Run()
        {
            // ExStart:CustomizingEmailHeader
            // Create an instance MailMessage class
            MailMessage message = new MailMessage();

            // Specify ReplyTo, From, To field
            message.ReplyToList.Add("reply@reply.com");
            message.From = "sender@sender.com";
            message.To.Add("receiver1@receiver.com");

            // Adding CC and BCC Addresses
            message.CC.Add("receiver2@receiver.com");
            message.Bcc.Add("receiver3@receiver.com");

            // Specify Message subject, Specify Date and Specify XMailer
            message.Subject = "test mail";
            message.Date = new DateTime(2006, 3, 6);
            message.XMailer = "Aspose.Email";
            message.Headers.Add_("secret-header", "mystery");

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
            Console.WriteLine(Environment.NewLine + "Email sent with customized headers.");
        }
    }
}
