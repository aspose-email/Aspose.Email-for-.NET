using System;
using System.Diagnostics;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class SettingHTMLBody
    {
        // ExStart:SettingHTMLBody
        public static void Run()
        {
            // Declare msg as MailMessage instance
            MailMessage msg = new MailMessage();

            // Use MailMessage properties like specify sender, recipient, message and HtmlBody
            msg.From = "newcustomeronnet@gmail.com";
            msg.To = "asposetest123@gmail.com";
            msg.Subject = "Test subject";
            msg.HtmlBody = "<html><body>This is the HTML body</body></html>";
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

            Console.WriteLine(Environment.NewLine + "Email sent with HTML body.");
        }

        private static SmtpClient GetSmtpClient()
        {
            SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password");
            client.SecurityOptions = SecurityOptions.Auto;
            return client;
        }
        // ExEnd:SettingHTMLBody
    }
}
