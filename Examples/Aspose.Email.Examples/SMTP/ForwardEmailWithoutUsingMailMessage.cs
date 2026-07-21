using Aspose.Email;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.SMTP
{
    class ForwardEmailWithoutUsingMailMessage
    {
        public static void Run()
        {
            string host = "mail.server.com";
            string username = "username";
            string password = "password";
            int smtpPort = 587;
            string sender = "Sender@domain.com";
            MailAddressCollection recipients = new MailAddressCollection();
            recipients.Add("recepient1@domain.com, recepient2@domain.com");

            using (SmtpClient client = new SmtpClient(host, smtpPort, username, password, SecurityOptions.Auto))
            {
                string fileName = @"test.eml";
                using (FileStream fs = File.OpenRead(Data.Email + fileName))
                {
                    client.Forward(sender, recipients, fs);
                }
            }
        }
    }
}
